#!/usr/bin/env python3
# coding: utf-8

"""
MRCC out file parser
This script allows the user to parse some information of MRCC out files and collect the results
in an xlsx file.
This script accepts a list of out files (.out).
The results are written in an xlsx file or to the stdout if no xlsx file is given.
Pre-requisite :
xlsxwriter package to be installed within the Python runtime environment
"""

import argparse
from pathlib import Path
import re

import xlsxwriter

parser = argparse.ArgumentParser()
parser.add_argument(
    "outfiles",
    nargs="+",
    type=argparse.FileType("rt"),
    default=sys.stdin,
    help="Path to one or more MRCC output files",
)
parser.add_argument(
    "--xlsxfile",
    nargs="?",
    type=argparse.FileType("wb"),
    default=sys.stdout,
    help="Path to result xlsx file",
)
args = parser.parse_args()
fpaths = [Path(outfile.name).resolve() for outfile in args.outfiles]

header0 = [
    "Molecule",
    "Method",
    "Basis",
    "CPU time",
    "Memory",
    "Disk load",
    "Energy",
    "Determinants",
]
header1 = [None, None, None, "[s]", "[MiB]", "[MiB]", "[A.U.]", None]
headers = (header0, header1)


def find(*, kw: str, text: str) -> str:
    """
    Find the value of a keyword in an MRCC output text

    :param kw: the keyword whose value we are looking for
    :param text: text of the MRCC out file
    :returns: the text value of keyword found first
    :raises KeyError: if the keyword is not found
    """
    search = re.search(rf"^\s?{kw}\s?=\s?([^#\n\r\s]+)", text, flags=re.IGNORECASE | re.MULTILINE)
    if not search:
        raise KeyError(kw)
    return search.group(1)


def get_energy(*, text: str) -> float:
    """
    Find the last "Total energy" value in an MRCC output text

    :param text: text of the MRCC out file
    :returns: the float value of the last "Total energy"
    :raises KeyError: if no "Total energy" is found
    """
    search = re.findall(
        r"^\s?Total\s[\w\d)(\]\[]+\senergy\s\[au\]:\s+(-?\d+\.\d*)",
        text,
        flags=re.IGNORECASE | re.MULTILINE,
    )
    if len(search) < 1:
        raise KeyError("Total energy")
    return float(search[-1])


def get_num_determinants(*, text: str) -> int:
    """
    Find the last "Total number of determinants" in an MRCC output text

    :param text: text of the MRCC out file
    :returns: the integer value of the last "Total number of determinants"
    :raises KeyError: if no "Total number of determinants" is found
    """
    search = re.findall(
        r"^\s?Total number of determinants:\s+(\d+)",
        text,
        flags=re.IGNORECASE | re.MULTILINE,
    )
    if len(search) < 1:
        raise KeyError("Total number of determinants")
    return int(search[-1])


workbook = xlsxwriter.Workbook(args.xlsxfile)
worksheet = workbook.add_worksheet()

bold = workbook.add_format({"bold": True})
worksheet.set_column("B:B", 20)

for header_num, header in enumerate(headers):
    worksheet.write_row(header_num, 0, header, bold)

for fnum, fpath in enumerate(fpaths):
    row_num = len(headers) + fnum
    molecule = fpath.stem.split("_")[0]

    with open(fpath, "rt") as out_file:
        out_text = out_file.read()

    calc = find(kw="calc", text=out_text).upper()
    basis = find(kw="basis", text=out_text).upper()
    num_determinants = get_num_determinants(text=out_text)
    energy = get_energy(text=out_text)

    worksheet.write(row_num, 0, molecule)
    worksheet.write(row_num, 1, calc)
    worksheet.write(row_num, 2, basis)
    worksheet.write(row_num, 6, energy)
    worksheet.write(row_num, 7, num_determinants)

workbook.close()
