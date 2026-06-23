from argparse import ArgumentParser
from datetime import datetime, timezone
from enum import IntEnum
from functools import lru_cache
from json import dump, load
from pathlib import Path
from re import fullmatch, search
from sys import exit, stderr
from time import gmtime, strftime, time
from xml.parsers.expat import ParserCreate
from zipfile import ZipFile

from pyarrow import csv, parquet


@lru_cache
def col_index(col):
    """Convert column letters like 'AZ' to a 0-based index."""
    result = 0
    for c in col:
        result = result * 26 + (ord(c) - ord("A") + 1)
    return result - 1

def cell_index(ref):
    """Convert a cell reference like 'AZ16' to (row, col) 0-based indices."""
    m = fullmatch(r"([A-Z]+)(\d+)", ref)
    return int(m.group(2)) - 1, col_index(m.group(1))

def index_col(index):
    """Convert a 0-based index to Excel-style column letters."""
    result = ""
    index += 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


warning_counts = {}

def warn(general, specific=""):
    """Print warning, suppressing after 20 occurrences of the same type."""
    count = warning_counts.get(general, 0)
    if count < 20:
        print(f"  Warning: {general} {specific}")
    elif count == 20:
        print(f"  Warning: Further '{general}' warnings suppressed.")
    warning_counts[general] = count + 1


class DT(IntEnum):
    """Excel cell datatypes, value reflects preference"""
    number   = 1
    boolean  = 2
    date     = 3
    time     = 4
    datetime = 5
    string   = 6
    formula  = 7
    error    = 8
    empty    = 9
    missing  = 10

# for each data type, as which other data types it can be read
COMPAT = {
    DT.number:   {DT.number, DT.string},
    DT.boolean:  {DT.boolean, DT.string},
    DT.date:     {DT.date, DT.datetime, DT.string},
    DT.time:     {DT.time, DT.string},
    DT.datetime: {DT.datetime, DT.string},
    DT.string:   {DT.string},
    DT.formula:  {DT.string},
    DT.error:    {DT.string},
    DT.empty:    {DT.number, DT.boolean, DT.date, DT.time, DT.datetime,
                  DT.string},
    DT.missing:  {DT.number, DT.boolean, DT.date, DT.time, DT.datetime,
                  DT.string}
}

# for each data type, the corresponding Arrow data type
ARROW = {
    DT.number:   "double",
    DT.boolean:  "bool",
    DT.date:     "date32[day]",
    DT.time:     "time64[ns]",
    DT.datetime: "timestamp[us]",
    DT.string:   "string",
    DT.missing:  "null"
}


def read_workbook(zf):
    """Return a list of sheet names and the epoch for dates."""
    sheet_names = []
    epoch = datetime(1899, 12, 30, tzinfo=timezone.utc)

    def start_element(name, attrs):
        nonlocal epoch
        if name == "sheet":
            sheet_names.append(attrs["name"])
        if (name == "workbookPr"
            and attrs.get("date1904", "0").lower() in ("1", "true")):
                epoch = datetime(1904, 1, 1, tzinfo=timezone.utc)

    p = ParserCreate()
    p.StartElementHandler = start_element
    p.ParseFile(zf.open("xl/workbook.xml"))
    return (sheet_names, epoch)


def read_shared_strings(zf):
    """Parse xl/sharedStrings.xml and return list of shared string values."""

    shared_strings = []
    csd = None              # current si contents
    ccd = ""                # current character data

    def start_element(name, attrs):
        nonlocal csd, ccd
        match name:
            case "si":
                csd = ""
            case "t":
                ccd = ""

    def end_element(name):
        nonlocal csd
        match name:
            case "si":
                shared_strings.append(csd)
            case "t":
                csd += ccd

    def character_data(data):
        nonlocal ccd
        ccd += data

    p = ParserCreate()
    p.StartElementHandler  = start_element
    p.EndElementHandler    = end_element
    p.CharacterDataHandler = character_data
    try:
        p.ParseFile(zf.open("xl/sharedStrings.xml"))
    except KeyError:
        warn("Unable to parse 'xl/sharedStrings.xml'.")
        return []

    return shared_strings


def classify_format(format_code):
    """Classify an Excel format code as 'number', 'date', 'time',
    or 'datetime'."""
    format_code = format_code.lower()
    # Check for date components (respecting m context)
    has_year = bool(search(r"(?<!y)y{2,4}(?!y)", format_code))
    has_day = bool(search(r"(?<![a-z])d{1,2}(?![a-z])", format_code))
    # Month: m not in minute context
    has_month = (search(r"(?<![a-z])m{1,2}(?![a-z])", format_code)
                 and not search(r"[h\]]:m|m:[s]", format_code))
    has_date = has_year or has_month or has_day
    # Check for clock time (h or hh, not [h])
    has_time = bool(search(r"(?<!\[)h{1,2}(?!\])", format_code))
    # combine
    if has_date and has_time:
        return "datetime"
    elif has_date:
        return "date"
    elif has_time:
        return "time"
    else:
        return "number"


BUILTIN_DATE_IDS = set(range(14, 18)) | set(range(27, 37)) | set(range(50, 59))
BUILTIN_TIME_IDS = set(range(18, 22))
BUILTIN_DATETIME_IDS = {22}

def read_number_types(zf):
    """Return a list which maps from cells' s attributes to number types."""

    number_formats = {}
    in_cellXfs = False
    number_types = []

    def start_element(name, attrs):
        nonlocal number_formats, in_cellXfs, number_types
        match name:
            case "numFmt":
                id = int(attrs["numFmtId"])
                number_formats[id] = attrs["formatCode"]
            case "cellXfs":
                in_cellXfs = True
            case "xf":
                if in_cellXfs:
                    id = int(attrs["numFmtId"])
                    if id in BUILTIN_DATE_IDS:
                        number_type = "date"
                    elif id in BUILTIN_TIME_IDS:
                        number_type = "time"
                    elif id in BUILTIN_DATETIME_IDS:
                        number_type = "datetime"
                    elif id in number_formats:
                        number_type = classify_format(number_formats[id])
                    else:
                        number_type = "number"
                    number_types.append(number_type)

    p = ParserCreate()
    p.StartElementHandler  = start_element
    try:
        p.ParseFile(zf.open("xl/styles.xml"))
    except KeyError:
        warn("Unable to parse 'xl/styles.xml'.")
        return []

    return number_types


def process_sheet(zf, sheet_index, cf, shared_strings, number_types,
                  epoch=None, prefer_formula=True, debug=False):
    """Parse worksheet XML, write CSV rows."""
    sri = None              # start row index
    sci = None              # start column index
    eri = None              # end row index
    eci = None              # end column index
    crc = None              # current row contents
    crt = None              # current row data types
    dtc = None              # data type counts per column
    crr = None              # current row 'r' attribute
    cri = None              # current row index
    ccr = None              # current cell 'r' attribute
    cci = None              # current column index
    orc = None              # output row and column
    cct = None              # current cell 't' attribute
    ccs = None              # current cell 's' attribute
    cvc = None              # current v contents
    cfc = None              # current f contents
    cic = None              # current is contents
    ccd = ""                # current character data
    ori = 0                 # output row index
    frc = None              # first row contents
    frt = None              # first row data types
    level = 0               # indentation for debug output
    lastrep = time()        # last time of progress report

    def start_element(name, attrs):
        nonlocal level
        nonlocal sri, sci, eri, eci
        nonlocal crc, crt, dtc, crr, cri
        nonlocal ccr, cci, orc, cct, ccs, cvc, cfc, cic, ccd
        if debug:
            astr = [f"{k}='{attrs[k]}'" for k in attrs]
            print(f"{'  ' * level}{name} {', '.join(astr)}", file=stderr)
        match name:
            case "dimension":
                dim = attrs["ref"]
                start, end = dim.split(":")
                sri, sci = cell_index(start)
                eri, eci = cell_index(end)
                print(f"  Range {dim}, {eri - sri + 1} × {eci - sci + 1}")
                dtc = [{} for _ in range(eci - sci + 1)]
            case "row":
                crc = [""] * (eci - sci + 1)
                crt = [DT.missing] * (eci - sci + 1)
                crr = attrs["r"]
                cri = int(crr) - 1
            case "c":
                ccr = attrs["r"]
                cci = col_index(ccr.removesuffix(crr))
                orc = f"({ori + 1}, {cci - sci + 1})"
                cct = attrs.get("t", "n")
                ccs = attrs.get("s")
                if ccs is not None and int(ccs) >= len(number_types):
                    warn(f"Reference to undefined style {ccs} in {ccr} {orc}")
                    ccs = None
                cvc = None
                cfc = None
                cic = None
            case "v":
                ccd = ""
            case "f":
                if (prefer_formula
                    and "ref" in attrs
                    and ":" in attrs["ref"]
                    and attrs["ref"].split(":")[1] != ccr):
                    warn(
                        "Array formula",
                        f"in {ccr} {orc} spills over {attrs['ref']}."
                    )
                ccd = ""
            case "is":
                cic = ""
            case "t":
                ccd = ""
        level += 1

    def end_element(name):
        nonlocal level, cvc, cfc, cic, ori, frc, frt, lastrep
        match name:
            case "row":
                # write row to csv
                print(",".join(crc), file=cf)
                if ori == 0:
                    # first row: remember
                    frc = crc
                    frt = crt
                else:
                    # other rows: count datatypes
                    for ci, dt in enumerate(crt):
                        dtc[ci][dt] = dtc[ci].get(dt, 0) + 1
                # update row number
                ori += 1
                # progress report
                if time() - lastrep > 0.05:
                    print(f"  {ori}", end="\r")
                    lastrep = time()
            case "c":
                if prefer_formula and cfc is not None:
                    warn(
                        f"Using formula text '{cfc}'",
                        f"as value in {ccr} {orc}"
                    )
                    crc[cci - sci] = '"' + cfc.replace('"', '""') + '"'
                    crt[cci - sci] = DT.formula
                else: # not prefer_formula or cfc is None:
                    dt, crc[cci - sci] = format_cell()
                    crt[cci - sci] = dt
                    if dt == DT.error:
                        warn(f"Formula with error '{cvc}'", f"in {ccr} {orc}")
            case "v":
                cvc = ccd
                if debug:
                    print(f"{'  ' * level}'{ccd}'", file=stderr)
            case "f":
                cfc = ccd
                if debug:
                    print(f"{'  ' * level}'{ccd}'", file=stderr)
            case "t":
                cic += ccd
                if debug:
                    print(f"{'  ' * level}'{ccd}'", file=stderr)
        level -= 1

    def character_data(data):
        nonlocal ccd
        ccd += data

    def format_cell():
        # Convert an Excel cell value to a CSV-safe string
        # based on cell type and formatting.
        match cct:
            case "inlineStr":
                # inline string
                # requires <is> to contain the string
                return DT.string, '"' + cic.replace('"', '""') + '"'
            case "str":
                # string produced by a formula
                # requires <v> to contain the string
                return DT.string, '"' + cvc.replace('"', '""') + '"'
            case "s":
                # shared string
                # requires <v> to contain the index into shared_strings
                str = shared_strings[int(cvc)]
                return DT.string, '"' + str.replace('"', '""') + '"'
            case "b":
                # boolean
                # requires <v> to be 0 or 1
                return DT.boolean, "TRUE" if cvc == "1" else "FALSE"
            case "e":
                # error in formula
                # requires <v> to contain the error code
                return DT.error, cvc
            case "d":
                # strict date or time
                # requires <v> to contain an ISO 8601 string
                # hh:mm:ss, YYYY-MM-DD, or YYYY-MM-DDThh:mm:ss
                if "-" in cvc and ":" in cvc:
                    dt = DT.datetime
                elif "-" in cvc:
                    dt = DT.date
                elif ":" in cvc:
                    dt = DT.time
                else:
                    warn(
                        "Unexpected date/time format",
                        f"'{cvc}' in {ccr}{orc}."
                    )
                    dt = DT.datetime      # fallback
                return dt, cvc
            case "n":
                # numeric: number, date, time, or datetime
                # requires <v> to contain the number
                # attribute s specifies formatting and thereby number type
                if cvc is None:
                    return DT.empty, ""
                if ccs is None:                                  # noqa: SIM108
                    # default number type
                    nt = "number"
                else:
                    # look up number type
                    nt = number_types[int(ccs)]
                # number
                if nt == "number":
                    return DT.number, cvc
                # serial date-times: days since epoch, fractional part is time
                value = float(cvc)
                # separate date and time components
                date = int(value)
                time = value - date
                # promote number type if necessary
                epsilon = value * 2**-53    # precision of value
                if ((nt == "date" and time > epsilon)
                    or (nt == "time" and date != 0)):
                    nt = "datetime"
                # format output
                match nt:
                    case "date":
                        # format date component
                        date = strftime(
                            "%Y-%m-%d",
                            gmtime(date * 86400 + epoch.timestamp())
                        )
                        return DT.date, date
                    case "time":
                        # format time component (ns)
                        # float64 precision at 24:00 is 0.0095 ns
                        time = round(time * 86400 * 1_000_000_000)
                        s, ns = divmod(time, 1_000_000_000)
                        m, s = divmod(s, 60)
                        h, m = divmod(m, 60)
                        if ns == 0:
                            time = f"{h:02d}:{m:02d}:{s:02d}"
                        else:
                            time = f"{h:02d}:{m:02d}:{s:02d}.{ns:09d}"
                            time = time.rstrip("0")
                        return DT.time, time
                    case "datetime":
                        # format date component
                        date = strftime(
                            "%Y-%m-%d",
                            gmtime(date * 86400 + epoch.timestamp())
                        )
                        # format time component (µs)
                        # float64 precision at 2026-01-01 is 0.441 µs
                        time = round(time * 86400 * 1_000_000)
                        s, us = divmod(time, 1_000_000)
                        m, s = divmod(s, 60)
                        h, m = divmod(m, 60)
                        if us == 0:
                            time = f"{h:02d}:{m:02d}:{s:02d}"
                        else:
                            time = f"{h:02d}:{m:02d}:{s:02d}.{us:06d}"
                            time = time.rstrip("0")
                        return DT.datetime, date + "T" + time
                    case _:
                        raise RuntimeError(
                            f"Unexpected number type attribute {nt}"
                        )
            case _:
                raise RuntimeError(f"Unexpected cell t attribute {cct}")

    start = time()
    p = ParserCreate()
    p.StartElementHandler  = start_element
    p.EndElementHandler    = end_element
    p.CharacterDataHandler = character_data
    p.ParseFile(zf.open(f"xl/worksheets/sheet{sheet_index}.xml"))
    stop = time()

    print(f"Wrote {ori:,} rows "
          f"in {stop - start:,.3g}s; "
          f"{eri - sri + 1 - ori:,} rows empty")

    return frc, frt, dtc


def guess_schema(first_row_contents, first_row_types, datatype_counts):
    """Infer column names and determine common data type per column."""
    print()
    print("Guessing schema")

    # is first row interpretable as column names?
    is_first_row_cols = all(ht in [DT.string, DT.empty, DT.missing]
                             for ht in first_row_types)
    print("First row: ", end="")
    print("column names" if is_first_row_cols else "data")

    # if not, update statistics with first row data types
    if not is_first_row_cols:
        for ci, dt in enumerate(first_row_types):
            datatype_counts[ci][dt] = datatype_counts[ci].get(dt, 0) + 1

    # number of columns
    num_cols = len(first_row_types)

    # determine column names
    if is_first_row_cols:
        # use first row for column names
        col_names = []
        for cn in first_row_contents:
            if len(cn) >= 2 and cn.startswith('"') and cn.endswith('"'):
                cn = cn[1:-1].replace('""', '"')
            col_names.append(cn)
    else:
        # empty names, to be fixed after
        col_names = [""] * num_cols
    # replace empty column names with letters
    col_names = [cn if len(cn) > 0 else index_col(ci)
                 for ci, cn in enumerate(col_names)]
    # ensure uniqueness
    col_set = set()
    for ci, cn in enumerate(col_names):
        while cn in col_set:
            cn += "_"
        col_names[ci] = cn
        col_set.add(cn)

    # for each column, determine best common data type
    common_datatype = [
        min(set.intersection(*(
                COMPAT[key] for key in datatype_counts[ci]
        )))
        for ci in range(num_cols)
    ]

    # for each column, show data type counts and best common data type
    print("Column data types")
    datatypes = set().union(*datatype_counts)
    # table header
    cnl = max(4, max(len(cn) for cn in col_names))
    print(f"  {'name':{cnl}}", end="")
    for dt in sorted(datatypes):
        print(f" {dt.name:>8}", end="")
    print(f"  {'common':{cnl}}")
    # counts
    for ci in range(num_cols):
        print(f"  {col_names[ci]:{cnl}}", end="")
        for dt in sorted(datatypes):
            if dt in datatype_counts[ci]:
                print(f" {datatype_counts[ci][dt]:8}", end="")
            else:
                print(f" {'·':>8}", end="")
        print(" ", common_datatype[ci].name)

    return is_first_row_cols, col_names, common_datatype


def process_excel(excel_path, sheet_name, csv_path, json_path,
                  prefer_formula=True, debug=False):
    """Parse xlsx file and write CSV + JSON schema options to disk."""
    with ZipFile(excel_path) as excel_file:
        # get sheet names
        sheet_names, epoch = read_workbook(excel_file)
        # get shared strings
        shared_strings = read_shared_strings(excel_file)
        # get number formats
        number_types = read_number_types(excel_file)

        # special syntax: list sheets
        if sheet_name is None:
            print()
            print("Select a sheet to process:")
            for sn in sheet_names:
                print(f"  '{sn}'")
            print()
            exit(0)

        # get sheet index
        try:
            sheet_index = sheet_names.index(sheet_name) + 1
        except ValueError:
            print()
            print(f"Error: Sheet '{sheet_name}' does not exist.")
            print()
            exit(2)

        # process sheet
        print()
        print(f"Reading Excel file '{excel_path}', sheet '{sheet_name}'")
        print(f"Writing CSV file '{csv_path}'")
        with open(csv_path, "w", encoding="utf-8", newline="\n") as csv_file:
            # read Excel, write CSV, and collect information
            (
                first_row_contents,
                first_row_types,
                datatype_counts
            ) = process_sheet(
                excel_file,
                sheet_index,
                csv_file,
                shared_strings=shared_strings,
                number_types=number_types,
                epoch=epoch,
                prefer_formula=prefer_formula,
                debug=debug
            )
    # guess schema from collected information
    (
        is_first_row_cols,
        col_names,
        common_datatype
    ) = guess_schema(
        first_row_contents,
        first_row_types,
        datatype_counts
    )
    # create options for reading CSV and save them as JSON
    csv_options = {
        "read_options": {
            "column_names": col_names,
            "skip_rows": 1 if is_first_row_cols else 0,
            # skip first line if it contains column names
            "block_size": 512 * 1024 * 1024
            # 512 MiB, modify if necessary
        },
        "parse_options": {
            "newlines_in_values": True
            # TODO: slow? fix by escaping newlines in CSV
        },
        "convert_options": {
            "column_types": {
                col_names[ci]: ARROW[common_datatype[ci]]
                for ci in range(len(col_names))
            },
            "true_values": ["TRUE"],
            "false_values": ["FALSE"],
            # only `TRUE` and `FALSE` are boolean
            "null_values": [""],
            # only `` is null
            "strings_can_be_null": True,
            # `` is null also in strings
            "quoted_strings_can_be_null": False
            # but `""` is empty string
        }
    }
    print(f"Writing JSON options '{json_path}'")
    with open(json_path, "w", encoding="utf-8", newline="\n") as json_file:
        dump(csv_options, json_file, indent=4, ensure_ascii=False)


def process_csv(csv_path, json_path, parquet_path):
    """Read CSV with PyArrow schema from JSON options and write Parquet."""
    print()
    print(f"Reading CSV file '{csv_path}' using JSON options '{json_path}'")
    print(f"Writing Parquet file '{parquet_path}'")
    # read options for csv.open_csv from JSON file
    with open(json_path, encoding="utf-8", newline="\n") as json_file:
        csv_options = load(json_file)
    with csv.open_csv(                                           # noqa: SIM117
        csv_path,
        csv.ReadOptions(**csv_options["read_options"]),
        csv.ParseOptions(**csv_options["parse_options"]),
        csv.ConvertOptions(**csv_options["convert_options"])
    ) as reader:
        with parquet.ParquetWriter(parquet_path, reader.schema) as writer:
            for bi, batch in enumerate(reader):
                writer.write_batch(batch)
                print(f"  {bi + 1}", end="\r")
    print(f"Wrote {bi + 1} batches")


def process(excel_filename, sheet_name, prefer_formula=False, debug=False):
    """Orchestrate the full pipeline: Excel → CSV → Parquet."""
    print()
    print("Salvaging Excel data")

    excel_path = Path(excel_filename)
    csv_path = Path(f"{excel_path.stem}-{sheet_name}.csv")
    json_path = Path(f"{csv_path.stem}.json")
    parquet_path = Path(csv_path.stem + ".parquet")


    if not excel_path.exists():
        print()
        print(f"Error: Excel file '{excel_path}' does not exist.")
        print()
        exit(1)

    if not csv_path.exists():
        # Idempotency: if CSV exists, skip Excel parsing. User can edit CSV
        # and rerun to regenerate Parquet without re-parsing the Excel file.
        if json_path.exists():
            print()
            print(f"Error: CSV file '{csv_path}' does not exist, "
                  f"but JSON options '{json_path}' do.")
            print("Not overwriting JSON options.")
            print()
            exit(3)
        process_excel(
            excel_path,
            sheet_name,
            csv_path,
            json_path,
            prefer_formula=prefer_formula,
            debug=debug
        )
    else:
        print()
        print(f"CSV file '{csv_path}' exists.")
        print(f"Not processing Excel file '{excel_path}'.")
        if excel_path.stat().st_mtime > csv_path.stat().st_mtime:
            print()
            print("Warning: Excel file is newer than CSV file.")

    if not json_path.exists():
        print()
        print(f"Error: CSV file '{csv_path}' exists, "
                f"but JSON options '{json_path}' do not.")
        print()
        exit(4)

    process_csv(csv_path, json_path, parquet_path)
    print()


def main():
    """CLI entry point."""
    parser = ArgumentParser(
        description="Salvage Excel data"
    )
    parser.add_argument(
        "excel_file",
        help="pathname of .xlsx file"
    )
    parser.add_argument(
        "sheet",
        nargs="?",
        help="sheet name"
    )
    parser.add_argument(
        "--prefer-cached",
        action="store_true",
        help="use cached cell values instead of formula text"
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="show debug information"
    )

    args = parser.parse_args()

    try:
        process(
            args.excel_file,
            args.sheet,
            prefer_formula=not args.prefer_cached,
            debug=args.debug
        )
    except KeyboardInterrupt:
        print()
        print("Aborted")
        print()
