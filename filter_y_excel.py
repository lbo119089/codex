import pandas as pd


def filter_rows_with_y_column(input_file: str, output_file: str, column_name: str) -> None:
    """Read an Excel file, filter rows where the specified column value is 'Y',
    and save the result to a new Excel file.

    Args:
        input_file: Path to the source Excel file.
        output_file: Path for the filtered Excel file.
        column_name: The name of the column to filter on.
    """
    df = pd.read_excel(input_file)
    filtered_df = df[df[column_name] == 'Y']
    filtered_df.to_excel(output_file, index=False)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Filter Excel rows where a column equals 'Y'.")
    parser.add_argument("input", help="Path to the input Excel file")
    parser.add_argument("output", help="Path for the output Excel file")
    parser.add_argument("--column", default="Status", help="Column name to filter on (default: 'Status')")
    args = parser.parse_args()

    filter_rows_with_y_column(args.input, args.output, args.column)

