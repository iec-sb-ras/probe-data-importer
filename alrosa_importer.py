import openpyxl

from rdflib import (
    Graph,
    Namespace,
    FOAF,
    XSD,
    RDF,
    RDFS,
    DCTERMS,
    URIRef,
    Literal,
    BNode,
)
from rdflib.namespace import WGS, SDO
import os.path
import unicodedata
from enum import Enum
import re
import requests as rq
from requests.auth import HTTPBasicAuth
import base64
import os
from namespace import PT, P, SCHEMA, BIBO, MT, GS, CGI, DBP, DBP_OWL
from pprint import pprint

import pudb

# Создание графа
G = Graph()

# Добавление алиасов (префиксов) для пространств имен
G.bind("foaf", FOAF)
G.bind("xsd", XSD)
G.bind("rdf", RDF)
G.bind("rdfs", RDFS)
G.bind("dcterms", DCTERMS)
G.bind("wgs", WGS)
G.bind("sdo", SDO)
G.bind("pt", PT)
G.bind("p", P)
G.bind("schema", SCHEMA)
G.bind("bibo", BIBO)
G.bind("mt", MT)
G.bind("gs", GS)
G.bind("cgi", CGI)
G.bind("dbp", DBP)
G.bind("dbp-owl", DBP_OWL)


def search_substring_on_sheet(sheet, column_number, substring):
    """Search for substring in a specific column of an openpyxl sheet."""
    matches = []
    for row in sheet.iter_rows(min_col=column_number, max_col=column_number):
        cell = row[0]
        if cell.value and substring.lower() in str(cell.value).lower():
            matches.append((cell.row, cell.value))
    return matches


def clean_value_of_cell(cell):
    """Extract and clean cell value, returning appropriate Python type."""
    value = cell.value

    if value is None:
        return None

    # Handle numeric types
    if isinstance(value, (int, float)):
        return value

    # Handle string values
    if isinstance(value, str):
        cleaned = value.strip()
        lc = cleaned.lower()

        if lc in ["н.д.", ""]:
            return None

        # Try to convert to int or float if possible
        try:
            # Check for integer
            if cleaned.isdigit():
                return int(cleaned)
            # Check for float
            cleaned = cleaned.replace(",", ".")
            return float(cleaned)
        except (ValueError, AttributeError):
            # Return cleaned string if conversion fails
            return cleaned

    # Return other types as-is (datetime, etc.)
    return value


def append_features(a_dict, name_row, value_row):
    """Append cleaned feature values from two rows to a dictionary."""
    if name_row and value_row:
        for name_cell, value_cell in zip(name_row, value_row):
            if name_cell.value and value_cell.value:
                key = str(name_cell.value).strip()
                val = clean_value_of_cell(value_cell)
                if val is not None:
                    a_dict[key] = val


def import_head_title_one_row_features(sheet, column_number, match_string) -> dict:
    """Extract features from the row containing match_string in the first column."""
    features = {}

    matches = search_substring_on_sheet(sheet, column_number, match_string)
    if not matches:
        return features  # Return empty dict if no match

    row_num, _ = matches[0]  # Use first match
    row = sheet[row_num]

    # Read next row for feature names
    next_row = sheet[row_num + 1] if row_num + 1 <= sheet.max_row else None
    # Read row after that for feature values
    value_row = sheet[row_num + 2] if row_num + 2 <= sheet.max_row else None

    append_features(features, next_row, value_row)

    return features

def import_head_title_multirow_as_data_frame(sheet, column_number, match_string):
    # Search analogous and start import as DF from next row as pandas DF


def add_to_dict_if_not_none(main_dict, name, a_dict):
    """Add a dictionary to main_dict under key 'name' if a_dict is not empty."""
    if a_dict:
        main_dict[name] = a_dict


def import_excel_table_into_dict(workbook, sheet_name_or_index):
    """
    Import data from Excel file into a dictionary using openpyxl.

    Args:
        file_path: Path to the Excel file
        sheet_name_or_index: Sheet name (str) or index (int)

    Returns:
        dict: Dictionary with row numbers as keys and row data as dicts
    """

    # Get the sheet
    if isinstance(sheet_name_or_index, str):
        sheet = workbook[sheet_name_or_index]
    else:
        sheet = workbook.worksheets[sheet_name_or_index]

    # print sheet Name
    print(f"Processing sheet: {sheet.title}")

    matches = search_substring_on_sheet(sheet, 0, "НДС")
    if matches:
        print("This sheet is not a tube")
        return None

    data_dict = {}

    features = {}

    tube_aim_features = import_head_title_one_row_features(
        sheet, 1, "Целевой показатель"
    )
    add_to_dict_if_not_none(features, "target", tube_aim_features)

    geology_aim_features = import_head_title_one_row_features(sheet, 1, "Геология")
    add_to_dict_if_not_none(features, "geology", geology_aim_features)

    olivine_aim_features = import_head_title_one_row_features(
        sheet, 1, "Оливины I-генерации"
    )
    add_to_dict_if_not_none(features, "olivine", olivine_aim_features)

    aa_aim_features = import_head_title_one_row_features(
        sheet, 1, "алмазная ассоциация"
    )
    add_to_dict_if_not_none(features, "assoc", aa_aim_features)

    add_to_dict_if_not_none(data_dict, "features", features)

    return data_dict


def search_excel_to_root(name):
    """
    Search for an Excel file by name in the current directory and its parent directories.

    Args:
        name (str): Name of the Excel file to search for.

    Returns:
        str or None: Full path to the Excel file if found, None otherwise.
    """
    current_dir = os.path.abspath(os.path.dirname(__file__))

    while True:
        # Check if file exists in current directory
        file_path = os.path.join(current_dir, name)
        if os.path.isfile(file_path):
            return file_path

        # Move to parent directory
        parent_dir = os.path.dirname(current_dir)
        if parent_dir == current_dir:  # Reached root directory
            break
        current_dir = parent_dir

    return None


def main():
    # Example usage

    file_path = "data/tubes.xlsx"
    workbook = openpyxl.load_workbook(search_excel_to_root(file_path), data_only=True)

    # for sheet_number in [1]:
    for sheet_number in range(len(workbook.sheetnames)):
        excel_data = import_excel_table_into_dict(workbook, sheet_number)
        pprint(excel_data)


if __name__ == "__main__":
    main()
