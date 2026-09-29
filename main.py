from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import StreamingResponse
import pandas as pd
import io
import re


app = FastAPI(title="Excel Data Merger")


# ---------------------------------------------------------
# STATIC + HTML
# ---------------------------------------------------------

app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {"request": request}
    )


# ---------------------------------------------------------
# BASIC HELPERS
# ---------------------------------------------------------

def clean_column_name(column):
    """
    Clean Excel column names.
    """
    if pd.isna(column):
        return ""

    column = str(column).strip()
    column = re.sub(r"\s+", " ", column)

    return column


def normalize_column_for_compare(column):
    """
    Used only when checking whether column names are common.
    Example:
        'Phone No' -> 'phone no'
        ' Phone  No ' -> 'phone no'
    """
    column = clean_column_name(column)

    return column.lower()


def normalize_matching_value(value):
    """
    Normalize the value used for matching.

    Examples:
        '  Ravi Kumar ' -> 'ravi kumar'
        '98765 43210' -> '9876543210'
    """

    if pd.isna(value):
        return ""

    value = str(value).strip().lower()

    # Remove Excel .0 from numeric-looking values
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]

    # Remove unnecessary spaces
    value = re.sub(r"\s+", " ", value)

    return value


def clean_excel_value(value):
    """
    Make values safe for Excel output.
    """

    if pd.isna(value):
        return ""

    # Avoid scientific notation for integer-like numbers
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))

    return value


def clean_dataframe(df):
    """
    Clean dataframe column names and cell values.
    """

    df = df.copy()

    df.columns = [clean_column_name(c) for c in df.columns]

    # Remove completely empty columns
    df = df.dropna(axis=1, how="all")

    # Clean values
    for column in df.columns:
        df[column] = df[column].map(clean_excel_value)

    return df


# ---------------------------------------------------------
# READ EXCEL FILES
# ---------------------------------------------------------

async def read_uploaded_files(files):
    """
    Read all uploaded Excel files.
    """

    if len(files) < 2:
        raise ValueError("Please upload at least 2 Excel files.")

    dataframes = []
    filenames = []

    for file in files:

        if not file.filename:
            raise ValueError("One of the uploaded files has no filename.")

        filename_lower = file.filename.lower()

        if not (
            filename_lower.endswith(".xlsx")
            or filename_lower.endswith(".xls")
        ):
            raise ValueError(
                f"{file.filename} is not an Excel file. "
                "Please upload .xlsx or .xls files."
            )

        content = await file.read()

        if not content:
            raise ValueError(
                f"{file.filename} is empty."
            )

        try:
            df = pd.read_excel(io.BytesIO(content))
        except Exception as e:
            raise ValueError(
                f"Could not read {file.filename}: {str(e)}"
            )

        df = clean_dataframe(df)

        if len(df.columns) == 0:
            raise ValueError(
                f"{file.filename} does not contain usable columns."
            )

        dataframes.append(df)
        filenames.append(file.filename)

    return dataframes, filenames


# ---------------------------------------------------------
# FIND COMMON COLUMNS
# ---------------------------------------------------------

def find_common_columns(dataframes):
    """
    Find columns that exist in ALL uploaded files.

    Comparison is case-insensitive and ignores extra spaces.
    """

    if not dataframes:
        return []

    # Map normalized name -> original column name
    first_map = {}

    for column in dataframes[0].columns:
        normalized = normalize_column_for_compare(column)

        if normalized:
            first_map[normalized] = column

    common_normalized = set(first_map.keys())

    for df in dataframes[1:]:

        current_columns = {
            normalize_column_for_compare(column)
            for column in df.columns
            if normalize_column_for_compare(column)
        }

        common_normalized &= current_columns

    common_columns = []

    for normalized in common_normalized:
        common_columns.append(first_map[normalized])

    # Sort for stable UI
    common_columns.sort(key=lambda x: x.lower())

    return common_columns


def find_real_column(df, requested_column):
    """
    Find the actual dataframe column corresponding
    to the selected matching column.
    """

    requested_normalized = normalize_column_for_compare(
        requested_column
    )

    for column in df.columns:

        if normalize_column_for_compare(column) == requested_normalized:
            return column

    return None


# ---------------------------------------------------------
# CREATE MERGE KEYS
# ---------------------------------------------------------

def add_merge_key(df, actual_column):
    """
    Add internal normalized matching key.
    """

    result = df.copy()

    result["_MERGE_KEY_"] = result[actual_column].map(
        normalize_matching_value
    )

    return result


# ---------------------------------------------------------
# VALIDATION
# ---------------------------------------------------------

def validate_records(
    dataframes,
    filenames,
    matching_column
):
    """
    Validate every record from every uploaded file.

    Categories:

    Blank:
        Matching value is empty.

    Duplicate:
        Same matching value occurs more than once
        inside the same file.

    Unmatched:
        Non-blank, non-duplicate value is missing
        from at least one other uploaded file.

    Matched:
        Non-blank, non-duplicate value exists
        in every uploaded file.
    """

    # -----------------------------------------------------
    # Find actual selected column in every file
    # -----------------------------------------------------

    actual_columns = []

    for df in dataframes:

        actual_column = find_real_column(
            df,
            matching_column
        )

        if actual_column is None:
            raise ValueError(
                f"Matching column '{matching_column}' "
                "was not found in all files."
            )

        actual_columns.append(actual_column)

    # -----------------------------------------------------
    # Get normalized values for each file
    # -----------------------------------------------------

    normalized_values_per_file = []

    for df, actual_column in zip(
        dataframes,
        actual_columns
    ):

        values = df[actual_column].map(
            normalize_matching_value
        )

        normalized_values_per_file.append(values)

    # -----------------------------------------------------
    # Count occurrence of each key in each file
    # -----------------------------------------------------

    occurrence_maps = []

    for values in normalized_values_per_file:

        counts = {}

        for value in values:

            if value == "":
                continue

            counts[value] = counts.get(value, 0) + 1

        occurrence_maps.append(counts)

    # -----------------------------------------------------
    # Determine which keys exist in all files
    # -----------------------------------------------------

    all_file_key_sets = []

    for counts in occurrence_maps:
        all_file_key_sets.append(set(counts.keys()))

    common_keys = set.intersection(
        *all_file_key_sets
    ) if all_file_key_sets else set()

    # -----------------------------------------------------
    # Validation containers
    # -----------------------------------------------------

    matched = []
    unmatched = []
    duplicate = []
    blank = []

    # -----------------------------------------------------
    # Validate every row
    # -----------------------------------------------------

    for file_index, (
        df,
        filename,
        actual_column,
        values,
        counts
    ) in enumerate(
        zip(
            dataframes,
            filenames,
            actual_columns,
            normalized_values_per_file,
            occurrence_maps
        )
    ):

        for row_index, value in enumerate(values):

            excel_row = row_index + 2

            original_value = df.iloc[
                row_index
            ][actual_column]

            # ---------------------------------------------
            # BLANK
            # ---------------------------------------------

            if value == "":

                blank.append({
                    "Source File": filename,
                    "Excel Row": excel_row,
                    "Matching Value": "",
                    "Reason": "Blank matching value"
                })

                continue

            # ---------------------------------------------
            # DUPLICATE
            # ---------------------------------------------

            if counts.get(value, 0) > 1:

                duplicate.append({
                    "Source File": filename,
                    "Excel Row": excel_row,
                    "Matching Value": original_value,
                    "Reason": "Matching value appears multiple times in this file"
                })

                continue

            # ---------------------------------------------
            # UNMATCHED
            # ---------------------------------------------

            if value not in common_keys:

                missing_files = []

                for other_index, other_counts in enumerate(
                    occurrence_maps
                ):

                    if value not in other_counts:

                        missing_files.append(
                            filenames[other_index]
                        )

                unmatched.append({
                    "Source File": filename,
                    "Excel Row": excel_row,
                    "Matching Value": original_value,
                    "Reason": "Matching value is missing from one or more files",
                    "Missing From": ", ".join(missing_files)
                })

                continue

            # ---------------------------------------------
            # MATCHED
            # ---------------------------------------------

            matched.append({
                "Source File": filename,
                "Excel Row": excel_row,
                "Matching Value": original_value,
                "Reason": "Valid match across all files"
            })

    return {
        "matched": matched,
        "unmatched": unmatched,
        "duplicate": duplicate,
        "blank": blank
    }


# ---------------------------------------------------------
# MERGE DATA
# ---------------------------------------------------------

def merge_dataframes(
    dataframes,
    matching_column
):
    """
    Merge all files using ONLY the selected matching column.

    The first file is the base file.

    Duplicate matching values in later files are reduced
    to the first occurrence so they do not multiply rows.
    """

    prepared = []

    actual_columns = []

    for df in dataframes:

        actual_column = find_real_column(
            df,
            matching_column
        )

        if actual_column is None:
            raise ValueError(
                f"Column '{matching_column}' "
                "does not exist in one of the files."
            )

        actual_columns.append(actual_column)

        temp = df.copy()

        temp["_MERGE_KEY_"] = temp[
            actual_column
        ].map(normalize_matching_value)

        prepared.append(temp)

    # -----------------------------------------------------
    # First file becomes base
    # -----------------------------------------------------

    result = prepared[0].copy()

    # Rename selected first column to requested name
    first_actual = actual_columns[0]

    if first_actual != matching_column:
        result = result.rename(
            columns={
                first_actual: matching_column
            }
        )

    # -----------------------------------------------------
    # Merge remaining files
    # -----------------------------------------------------

    for file_index in range(
        1,
        len(prepared)
    ):

        current = prepared[file_index].copy()

        current_actual = actual_columns[file_index]

        # Remove duplicate matching keys from current file
        current = current.drop_duplicates(
            subset=["_MERGE_KEY_"],
            keep="first"
        )

        # Rename columns to avoid collisions
        rename_map = {}

        for column in current.columns:

            if column == "_MERGE_KEY_":
                continue

            # Selected matching column is not added again
            if column == current_actual:
                continue

            if column in result.columns:

                rename_map[column] = (
                    f"{column}_file{file_index + 1}"
                )

        current = current.rename(
            columns=rename_map
        )

        # -------------------------------------------------
        # Remove duplicate selected column
        # -------------------------------------------------

        columns_to_keep = [
            column
            for column in current.columns
            if column != current_actual
        ]

        current = current[
            columns_to_keep
        ]

        # -------------------------------------------------
        # Merge
        # -------------------------------------------------

        result = result.merge(
            current,
            on="_MERGE_KEY_",
            how="left"
        )

    # -----------------------------------------------------
    # Remove internal key
    # -----------------------------------------------------

    if "_MERGE_KEY_" in result.columns:
        result = result.drop(
            columns=["_MERGE_KEY_"]
        )

    # -----------------------------------------------------
    # Clean output
    # -----------------------------------------------------

    result = result.fillna("")

    for column in result.columns:
        result[column] = result[column].map(
            clean_excel_value
        )

    return result


# ---------------------------------------------------------
# EXCEL OUTPUT
# ---------------------------------------------------------

def dataframe_to_excel_response(
    df,
    filename
):
    """
    Convert dataframe to downloadable Excel.
    """

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Data"
        )

    output.seek(0)

    return StreamingResponse(
        output,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition":
                f'attachment; filename="{filename}"'
        }
    )


# ---------------------------------------------------------
# VALIDATION DATAFRAME
# ---------------------------------------------------------

def validation_to_dataframe(
    records,
    dataframes,
    filenames,
    matching_column
):
    """
    Create detailed validation Excel data.

    Includes:
        Source File
        Excel Row
        Matching Value
        Reason
        Missing From
        Original row data
    """

    final_records = []

    for record in records:

        filename = record.get(
            "Source File",
            ""
        )

        excel_row = record.get(
            "Excel Row",
            ""
        )

        # Find source file
        try:
            file_index = filenames.index(
                filename
            )
        except ValueError:
            file_index = -1

        output_record = dict(record)

        # Add original row values
        if file_index >= 0:

            df = dataframes[file_index]

            dataframe_row_index = excel_row - 2

            if (
                dataframe_row_index >= 0
                and dataframe_row_index < len(df)
            ):

                original_row = df.iloc[
                    dataframe_row_index
                ]

                for column in df.columns:

                    # Do not overwrite validation columns
                    if column not in output_record:

                        output_record[column] = (
                            clean_excel_value(
                                original_row[column]
                            )
                        )

        final_records.append(
            output_record
        )

    if not final_records:

        return pd.DataFrame(
            columns=[
                "Source File",
                "Excel Row",
                "Matching Value",
                "Reason"
            ]
        )

    return pd.DataFrame(
        final_records
    )


# ---------------------------------------------------------
# ANALYZE
# ---------------------------------------------------------

@app.post("/analyze")
async def analyze(
    files: list[UploadFile] = File(...)
):

    try:

        dataframes, filenames = (
            await read_uploaded_files(files)
        )

        common_columns = find_common_columns(
            dataframes
        )

        file_information = []

        for df, filename in zip(
            dataframes,
            filenames
        ):

            file_information.append({
                "filename": filename,
                "rows": len(df),
                "columns": list(df.columns)
            })

        return {
            "success": True,
            "files": file_information,
            "common_columns": common_columns
        }

    except Exception as e:

        return {
            "success": False,
            "message": str(e)
        }


# ---------------------------------------------------------
# MERGE
# ---------------------------------------------------------

@app.post("/merge")
async def merge(
    files: list[UploadFile] = File(...),
    matching_column: str = Form(...)
):

    try:

        dataframes, filenames = (
            await read_uploaded_files(files)
        )

        if len(dataframes) < 2:

            return {
                "success": False,
                "message": "Please upload at least 2 files."
            }

        common_columns = find_common_columns(
            dataframes
        )

        # Verify selected column is actually common
        selected_normalized = (
            normalize_column_for_compare(
                matching_column
            )
        )

        common_normalized = {
            normalize_column_for_compare(column)
            for column in common_columns
        }

        if selected_normalized not in common_normalized:

            return {
                "success": False,
                "message": (
                    "The selected matching column "
                    "does not exist in all uploaded files."
                )
            }

        # Merge
        merged_df = merge_dataframes(
            dataframes,
            matching_column
        )

        # Validation
        validation = validate_records(
            dataframes,
            filenames,
            matching_column
        )

        return {
            "success": True,
            "rows": len(merged_df),
            "columns": list(merged_df.columns),
            "preview": merged_df.to_dict(
                orient="records"
            ),
            "validation": {
                "matched": len(
                    validation["matched"]
                ),
                "unmatched": len(
                    validation["unmatched"]
                ),
                "duplicate": len(
                    validation["duplicate"]
                ),
                "blank": len(
                    validation["blank"]
                )
            },
            "unmatched_records": validation[
                "unmatched"
            ],
            "duplicate_records": validation[
                "duplicate"
            ],
            "blank_records": validation[
                "blank"
            ]
        }

    except Exception as e:

        return {
            "success": False,
            "message": str(e)
        }


# ---------------------------------------------------------
# DOWNLOAD MERGED EXCEL
# ---------------------------------------------------------

@app.post("/download")
async def download(
    files: list[UploadFile] = File(...),
    matching_column: str = Form(...)
):

    try:

        dataframes, filenames = (
            await read_uploaded_files(files)
        )

        merged_df = merge_dataframes(
            dataframes,
            matching_column
        )

        return dataframe_to_excel_response(
            merged_df,
            "merged_output.xlsx"
        )

    except Exception as e:

        return {
            "success": False,
            "message": str(e)
        }


# ---------------------------------------------------------
# DOWNLOAD VALIDATION CATEGORY
# ---------------------------------------------------------

@app.post("/download-validation")
async def download_validation(
    files: list[UploadFile] = File(...),
    matching_column: str = Form(...),
    validation_type: str = Form(...)
):

    try:

        dataframes, filenames = (
            await read_uploaded_files(files)
        )

        validation = validate_records(
            dataframes,
            filenames,
            matching_column
        )

        valid_types = {
            "matched",
            "unmatched",
            "duplicate",
            "blank"
        }

        if validation_type not in valid_types:

            return {
                "success": False,
                "message": "Invalid validation type."
            }

        records = validation[
            validation_type
        ]

        result_df = validation_to_dataframe(
            records,
            dataframes,
            filenames,
            matching_column
        )

        output_filename = (
            f"{validation_type}_records.xlsx"
        )

        return dataframe_to_excel_response(
            result_df,
            output_filename
        )

    except Exception as e:

        return {
            "success": False,
            "message": str(e)
        }