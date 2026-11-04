#!/usr/bin/env python3

# /// script
# requires-python = ">=3.14"
# dependencies = [
#     "pyserial",
#     "openpyxl",
#     "pint",
#     "python-ranges",
# ]
# ///

import argparse
import functools
import itertools
import json
import logging
import math
import threading
import tkinter as tk
from abc import ABC, abstractmethod
from collections.abc import Iterable
from contextlib import AbstractContextManager
from dataclasses import dataclass, field
from pathlib import Path

import pint
import serial
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from ranges import Range, RangeSet

ureg = pint.UnitRegistry()
log = logging.getLogger(__name__)

@dataclass(frozen=True)
class Cell:
	ranges: RangeSet = field(default_factory=RangeSet)
	label: str = field(default="")

@dataclass(frozen=True, kw_only=True)
class Scheme(ABC):
	count: int
	label: str = ""

	@staticmethod
	def parse_range(start: float | None, end: float | None) -> Range:
		kw_args = {}
		if start is not None:
			kw_args["start"] = float(start)
		if end is not None:
			kw_args["end"] = float(end)
		return Range(**kw_args)

	@abstractmethod
	def ranges(self) -> Iterable[RangeSet]:
		""" Generates nominal ranges for single cell """
		...

	def generate(self) -> Iterable[Cell]:
		return (Cell(ranges=rs, label=self.label) for rs in self.ranges())


@dataclass(frozen=True, kw_only=True)
class Linear(Scheme):
	start: float | None = None
	end: float | None = None

	def ranges(self) -> Iterable[RangeSet]:
		r = Scheme.parse_range(self.start, self.end)
		if r.isinfinite():
			return (RangeSet(r) for _ in itertools.repeat(None, self.count))
		
		steps = itertools.accumulate(itertools.repeat(r.length() / self.count, self.count), initial=r.start)
		return (RangeSet(Range(s, e)) for s, e in itertools.pairwise(steps))


@dataclass(frozen=True, kw_only=True)
class Log(Scheme):
	
	start: float
	end: float
	k: float = 1.0

	def ranges(self) -> Iterable[RangeSet]:
		if abs(self.k) < 1e-9:
			return Linear(count=self.count, start=self.start, end=self.end, label=self.label).ranges()

		r = Scheme.parse_range(self.start, self.end)
		length = r.length()
		scale = 1.0 / (math.exp(self.k) - 1)
		
		def get_val(step):
			return self.start + length * (math.exp(self.k * (step / self.count)) - 1) * scale

		return (
			RangeSet(Range(get_val(s), get_val(s + 1))) 
			for s in range(self.count)
		)


@dataclass(frozen=True, kw_only=True)
class Empty(Scheme):
	def ranges(self) -> Iterable[RangeSet]:
		return (RangeSet() for _ in itertools.repeat(None, self.count))


def format_quantity(q: pint.Quantity) -> str:
	if math.isinf(q.magnitude):
		return "∞" if q.magnitude > 0 else "-∞"

	qc = q.to_compact()
	return f"{f"{qc.magnitude:.2f}".rstrip("0").rstrip(".")} {qc.units:~#P}"

class Scpi(AbstractContextManager["Scpi"]):
	
	port: serial.Serial
	
	def __init__(self, port: str, unit: pint.Unit):
		self.port = serial.Serial(port, 115200, timeout=0.5)
		self.port.reset_input_buffer()
		self.port.reset_output_buffer()
		self.unit = unit
		self.scale = 1 / ureg.get_base_units(unit)[0]

	def __exit__(self, *_):
		self.port.close()

	def read(self) -> pint.Quantity | None:
		self.port.write(b"FETC?\n")
		line = self.port.read_until(b"\r\n").decode(errors="replace").strip()
		if not line or line.lower().startswith("cmd err"):
			return None
		try:
			value = float(line.split(",")[0]) * self.scale
		except ValueError:
			log.debug(f"Can't parse nominal value: {line!r}", exc_info=True)
			return None
		return self.unit * value if math.isfinite(value) else None


@dataclass(frozen=True, init=False)
class Box:
	name: str
	unit: pint.Unit
	layout: list[int]
	cells: list[Cell]
	
	def __init__(
		self,
		name: str,
		unit: pint.Unit,
		layout: list[int],
		schemes: Iterable[Iterable[Cell]] = (),
		merges: Iterable[Iterable[int | list[int]]] = ()
	):
		object.__setattr__(self, "name", name)
		object.__setattr__(self, "unit", unit)
		object.__setattr__(self, "layout", layout)

		cells = list(
			itertools.islice(
				itertools.chain(
					itertools.chain.from_iterable(schemes),
					(Cell() for _ in itertools.repeat(None)),
				),
				math.prod(self.layout),
			)
		)

		for merge in merges:
			merged_ranges = RangeSet()
			for coord in merge:
				merged_ranges.add(cells[self.flat(coord)].ranges)
			for coord in merge:
				cell_ranges = cells[self.flat(coord)].ranges
				cell_ranges.clear()
				cell_ranges.update(merged_ranges)

		object.__setattr__(self, "cells", cells)

	def coords(self, index: int):
		result = [0] * len(self.layout)
		for i, dim in enumerate(reversed(self.layout)):
			index, result[-1 - i] = divmod(index, dim)
		return result

	def flat(self, coords: int | Iterable[int]) -> int:
		return (
			coords
			if isinstance(coords, int)
			else functools.reduce(
				lambda acc, element: acc * element[0] + element[1],
				zip(self.layout, coords),
				0,
			)
		)

	def cell(self, coords: int | list[int]):
		return self.cells[self.flat(coords)]

	def find(self, nominal: float) -> Iterable[int]:
		return ( index for index, cell in enumerate(self.cells) if nominal in cell.ranges )

	def format_cell(self, cell: Cell):
		value = "\n".join(
			itertools.chain(
				(
					f"{format_quantity(self.unit * r.start)} .. {format_quantity(self.unit * r.end)}"
					for r in cell.ranges
				),
				(cell.label,) if cell.label else (),
			)
		)
		return value if value else "(empty)"

	def export(self, path=None):
		if path is None:
			path = f"{self.name}.xlsx"
			
		dim = 20
		
		wb = Workbook()
		ws = wb.active
		ws.title = self.name

		cols, rows = self.layout[-1], math.prod(self.layout[:-1])

		for col in range(cols):
			ws.column_dimensions[get_column_letter(col + 1)].width = dim

		for row in range(rows):
			ws.row_dimensions[row + 1].height = dim * 5.15

		center_alignment = Alignment(horizontal="center", vertical="center")

		for index, cell in enumerate(self.cells):
			row = index // cols + 1
			col = index % cols + 1

			excel_cell = ws.cell(row, col, self.format_cell(cell))
			excel_cell.alignment = center_alignment

		wb.save(path)

	def scpi(self, port):
		with Scpi(port, self.unit) as s:
			print("[CONNECTED] Hit Enter to continue")
			while True:
				input()
				if (quantity := s.read()) is not None:
					bins = self.find(quantity.magnitude)
					print(f"Nominal: {format_quantity(quantity)}; Bins: {list(map(self.coords, bins))}", end="")
				else:
					print("Can't read nominal from device")

	def ui(self, port: str):
		root = tk.Tk()
		root.configure(bg="white")
		root.title("Storage Box")
		root.state('zoomed') if root.tk.call('tk', 'windowingsystem') == 'win32' else root.attributes('-zoomed', True)

		cols, rows = self.layout[-1], math.prod(self.layout[:-1])
		for index, cell in enumerate(self.cells):
			lbl = tk.Label(
				root,
				text=self.format_cell(cell),
				bg="white",
				relief="ridge",
				anchor="center",
				justify="center",
			)
			lbl.grid(row=index // cols, column=index % cols, sticky="nsew")

		for r in range(rows):
			root.rowconfigure(r, weight=1)
		for c in range(cols):
			root.columnconfigure(c, weight=1)

		def update_ui(q: pint.Quantity | None):
			labels = root.winfo_children()
			for idx in getattr(root, 'last_highlighted_indices', []):
				labels[idx].configure(bg="white")
			indices = list(self.find(q.magnitude)) if q is not None else []
			for idx in indices:
				labels[idx].configure(bg="lightgreen")
			root.last_highlighted_indices = indices
		
		stop = threading.Event()

		def worker(port):
			while not stop.is_set():
				try:
					with Scpi(port, self.unit) as scpi:
						while not stop.is_set():
							quantity = scpi.read()
							log.debug(f"Received nominal quantity: {quantity}")
							try:
								root.after(0, update_ui, quantity)
							except (tk.TclError, RuntimeError):
								return
							stop.wait(0.2)
				except serial.SerialException:
					log.exception(f"Serial port ({port}) error. Reconnecting in 5 seconds...")
					stop.wait(5)

		thread = threading.Thread(target=worker, args=(port,), daemon=True)
		thread.start()
		try:
			root.mainloop()
		finally:
			stop.set()
			thread.join(timeout=2)


	# Метод _parse_range больше не нужен, так как Range создается внутри классов-инициализаторов
	@staticmethod
	def from_json(path: Path) -> 'Box':
		with open(path, 'r', encoding="utf-8") as f:
			data: dict[str, object] = json.load(f)
	
		scheme_types = {cls.__name__.lower(): cls for cls in Scheme.__subclasses__()}

		def parse_scheme(scheme: dict[str, object]) -> Iterable[Cell]:
			cls = scheme_types[scheme["type"].lower()]
			params = {k: v for k, v in scheme.items() if k != "type"}
			return cls(**params).generate()

		box = Box(
			name = path.stem,
			unit = ureg.parse_units(data["unit"]),
			layout = data["layout"],
			schemes = ( parse_scheme(scheme) for scheme in data["schemes"] ),
			merges = data.get("merges", ())
		)

		return box


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description="Storage Box Utility")
	parser.add_argument("json_path", type=Path, help="Path to the JSON configuration file")
	parser.add_argument("command", choices=["ui", "scpi", "export"], help="Command to execute")
	parser.add_argument("--port", "-p", default="/dev/ttyUSB0", help="Serial port (default: /dev/ttyUSB0)")
	parser.add_argument("--output", "-o", default=None, help="Excel output path (default: {json_path.stem}.xlsx)")
	parser.add_argument("--debug", action="store_true", help="Enable debug logging")
	
	args = parser.parse_args()
	
	logging.basicConfig(level=logging.DEBUG if args.debug else logging.INFO)
	
	box = Box.from_json(args.json_path)
	
	match args.command:
		case "ui":
			box.ui(port=args.port)
		case "scpi":
			box.scpi(port=args.port)
		case "export":
			box.export(path=args.output)

