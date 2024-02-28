"""
This script compares two Excel files (one manual and one automated) and writes the results to a 
    text file and an Excel file.

The script performs the following steps:
1. Reads in the manual and automated files.
2. Removes 'Id' and 'VIA' columns from both files if they exist.
3. Performs basic validation on the two files, checking if they have the same number of rows 
    and columns, the same column names, and the same data types for each column.
4. If the files pass basic validation, it renames and drops columns in both files based on a mapping file.
5. Sorts the rows in both files based on specified columns.
6. Compares the sorted files and writes any differences to the text file and the Excel file.

The script can be configured to ignore differences in the number of rows and/or data types.

Parameters:
_vendor (str): The vendor name.
_filename (str): The filename.
_new_file (int): Whether to start a new results file or append to an existing one.
_ignore_rows (int): Whether to ignore the number of rows in the comparison.
_ignore_data_types (int): Whether to ignore the data types in the comparison.
_error_codes (dict): The dictionary of error codes.

Returns:
None
"""

from pathlib import Path
import pandas as pd
import numpy as np
from datetime import datetime


def rename_and_drop_columns(input_df, column_mapping_df):
    """
    Renames and drops columns in a DataFrame based on a mapping DataFrame.

    Parameters:
    input_df (pd.DataFrame): The DataFrame to modify.
    column_mapping_df (pd.DataFrame): The DataFrame containing the column mappings.

    Returns:
    pd.DataFrame: The modified DataFrame.
    """
    # Create a dictionary from the column_mapping dataframe
    column_mapping_dict = column_mapping_df.set_index(
        'file_column')['rpm_column'].to_dict()

    # Rename the columns in the manual dataframe
    input_df = input_df.rename(columns=column_mapping_dict)

    # Find the columns in the manual dataframe that are not in the column_mapping dataframe
    columns_to_drop = [
        col for col in input_df.columns if col not in column_mapping_dict.values()]

    # Drop the columns that are not in the column_mapping dataframe
    input_df = input_df.drop(columns=columns_to_drop)

    return input_df


def basic_conditioning(df1, df2):
    """
    Removes 'Id' and 'VIA' columns from two DataFrames if they exist.

    Parameters:
    df1 (pd.DataFrame): The first DataFrame.
    df2 (pd.DataFrame): The second DataFrame.

    Returns:
    tuple: The modified DataFrames.
    """
    # Remove 'Id' column and 'VIA' Column if any exists
    if any(column_name == 'id' for column_name in df2.columns):
        df2.drop(['id'], axis=1, inplace=True)
    if any(column_name == 'Id' for column_name in df2.columns):
        df2.drop(['id'], axis=1, inplace=True)
    if any(column_name == 'VIA' for column_name in df1.columns):
        df1.drop(['VIA'], axis=1, inplace=True)
    if any(column_name == 'Via' for column_name in df1.columns):
        df1.drop(['Via'], axis=1, inplace=True)
    if any(column_name == 'VIA' for column_name in df2.columns):
        df2.drop(['VIA'], axis=1, inplace=True)
    if any(column_name == 'Via' for column_name in df2.columns):
        df2.drop(['Via'], axis=1, inplace=True)
    if any(column_name == 'Original Provider name ' for column_name in df1.columns):
        df1.drop(['Original Provider name '], axis=1, inplace=True)
    return df1, df2


def compare_summaries(df1, df2):
    """
    Compares two DataFrames after deleting all columns except "Account" and "Gross commission",
    converting "Account" to string, grouping by "Account", aggregating "Gross commission" by sum, 
    and sorting by "Account".

    Parameters:
    df1 (pd.DataFrame): The first DataFrame.
    df2 (pd.DataFrame): The second DataFrame.

    Returns:
    tuple: A code indicating whether the DataFrames are the same (0) or different (1),
           and a DataFrame containing the differences (or an empty DataFrame if they are the same).
    """
    # Select only the "Account" and "Gross commission" columns, convert "Account" to string, group by "Account", sum "Gross commission", and sort by "Account"
    df1 = df1[['Account', 'Gross commission']]
    df1['Account'] = df1['Account'].astype(str)
    df1 = df1.groupby('Account').sum().sort_index()
    df1 = df1.reset_index(drop=False)

    df2 = df2[['Account', 'Gross commission']]
    df2['Account'] = df2['Account'].astype(str)
    df2 = df2.groupby('Account').sum().sort_index()
    df2 = df2.reset_index(drop=False)

    # Compare the two DataFrames
    if df1.equals(df2):
        # If they are the same, return 0 and an empty DataFrame
        return 0, pd.DataFrame()
    else:
        # If they are different, return 1 and a DataFrame with the differences
        diff = df1.compare(df2)
        return 1, diff


def basic_validation(df1, df2, result_file, _ignore_rows=0, _ignore_data_types=0):
    """
    Performs basic validation on two DataFrames and writes the results to a file.

    Parameters:
    df1 (pd.DataFrame): The first DataFrame.
    df2 (pd.DataFrame): The second DataFrame.
    result_file (file object): The file to write the results to.
    _ignore_rows (int, optional): Whether to ignore the number of rows in the comparison. Defaults to 0.
    _ignore_data_types (int, optional): Whether to ignore the data types in the comparison. Defaults to 0.

    Returns:
    tuple: The error code and a DataFrame or list containing the differences.
    """
    # Check if the dataframes have the same number of rows and columns
    if df1.shape == df2.shape:
        result_file.write(
            "The dataframes have the same number of rows and columns.\n")
        print("The dataframes have the same number of rows and columns.\n")
    else:
        result_file.write(
            "The dataframes do not have the same number of rows and columns.\n")
        result_file.write(
            f"The manual dataframe has {df1.shape[0]} rows and {df1.shape[1]} columns.\n")
        result_file.write(
            f"The automated dataframe has {df2.shape[0]} rows and {df2.shape[1]} columns.\n")
        print(
            f"The manual dataframe has {df1.shape[0]} rows and {df1.shape[1]} columns.\n")
        print(
            f"The automated dataframe has {df2.shape[0]} rows and {df2.shape[1]} columns.\n")
        # Check if the number of columns is the same.
        if df1.shape[1] == df2.shape[1]:
            if _ignore_rows == 1:
                result_file.write(
                    "Ignoring the number of rows in the comparison.\n")
                print("Ignoring the number of rows in the comparison.")
                error = 3
            else:
                error = 2
                return error, pd.DataFrame()
        else:
            error = 1
            return error, pd.DataFrame()

    # Check if the dataframes have the same columns
    if set(df1.columns) == set(df2.columns):
        # Check if corresponding columns in both dataframes have the same data type
        different_types = 0
        comparison_df = pd.DataFrame(
            columns=['Column', 'Same DataType', 'DataTypeManual', 'DataTypeAutomated'])
        for column in df1.columns:
            if df1[column].dtype != df2[column].dtype:
                result_file.write(
                    f"The column '{column}' has different data types in the dataframes.\n")
                print(
                    f"The column '{column}' in manual has type {df1[column].dtype}, while in automated has type {df2[column].dtype} \n")
                different_types = 1
                new_row = pd.DataFrame([{'Column': column, 'Same DataType': 'No',
                                         'DataTypeManual': df1[column].dtype,
                                         'DataTypeAutomated': df2[column].dtype}])
                comparison_df = pd.concat([comparison_df,
                                          new_row], ignore_index=True)

                # If the column have different data types, ask user if they want to copy the data type from one file to the other
                copy_type = input("Do you want to copy the data type from one file to the other? \n \
                                   no: enter 0 \n \
                                   manual to automated: enter 1 \n \
                                   automated to manual: enter 2 \n")
                match copy_type:
                    case '0':
                        print("No data type copied")
                    case '1':
                        df2[column] = df2[column].astype(df1[column].dtype)
                        result_file.write(
                            f"The data type of column '{column}' in the automated file has been changed to match the manual file.\n")
                        print(
                            f"The data type of column '{column}' in the automated file has been changed to match the manual file.")
                        _ignore_data_types = 1
                    case '2':
                        df1[column] = df1[column].astype(df2[column].dtype)
                        result_file.write(
                            f"The data type of column '{column}' in the manual file has been changed to match the autoated file.\n")
                        print(
                            f"The data type of column '{column}' in the manual file has been changed to match the automated file.")
                        _ignore_data_types = 1
        if different_types == 1:
            if _ignore_data_types == 1:
                result_file.write(
                    "Ignoring the data types in the comparison.\n")
                error = 9
            else:
                error = 5
            return error, comparison_df
        else:
            error = 6
            result_file.write(
                "The dataframes have the same columns and data types.\n")
            print("The dataframes have the same columns and data types.\n")
            return error, pd.DataFrame()
    else:
        error = 4
        different_columns = list(set(df1.columns) ^ set(df2.columns))
        result_file.write(
            f"The dataframes have different columns: {different_columns}\n")
        return error, pd.DataFrame(different_columns)


def sort_rows(unsorted_df, columns_to_sort):
    """
    Sorts a DataFrame based on specified columns.

    Parameters:
    unsorted_df (pd.DataFrame): The DataFrame to sort.
    columns_to_sort (tuple): The columns to sort by.

    Returns:
    tuple: The error code and the sorted DataFrame.
    """
    # Check if all columns in columns_to_sort exist in the dataframe
    if all(item in unsorted_df.columns for item in columns_to_sort):
        # Sort the dataframe
        df_sorted = unsorted_df.sort_values(
            list(columns_to_sort)).reset_index(drop=True)
        return (0, df_sorted)
    else:
        return (8, pd.DataFrame())


def compare_df(df1, df2, result_file):
    """
    Compares two DataFrames and writes the differences to a file.

    Parameters:
    df1 (pd.DataFrame): The first DataFrame.
    df2 (pd.DataFrame): The second DataFrame.
    result_file (file object): The file to write the differences to.

    Returns:
    tuple: The error code and a DataFrame containing the differences.
    """
    # Test if the total sum of the "Gross commission" column for each "Account" is the same in both dataframes
    summary_error, summary_diff = compare_summaries(df1, df2)
    if summary_error == 1:
        result_file.write(
            "The dataframes have different total sum of 'Gross commission' for each 'Account'.\n")
        print("The dataframes have different total sum of 'Gross commission' for each 'Account'.\n")
        summary_diff = summary_diff.reset_index().rename(
            columns={'self': 'manual', 'other': 'automated'})
        return 10, summary_diff
    else:
        result_file.write(
            "The dataframes have the same total sum of 'Gross commission' for each 'Account'.\n")
        print("The dataframes have the same total sum of 'Gross commission' for each 'Account'.\n")

    # Start cell by cell comparison

    # Select only numeric columns
    df1_numeric = df1.select_dtypes(
        include=[np.number]).reset_index(drop=True)
    df2_numeric = df2.select_dtypes(
        include=[np.number]).reset_index(drop=True)

    # Create a boolean mask for values in df1_numeric and df2_numeric that are not close
    mask = ~np.isclose(df1_numeric.values, df2_numeric.values, atol=0.01)

    # Apply the mask to df1_numeric and df2_numeric
    if df1_numeric.shape == df2_numeric.shape:
        diff_numeric = df1_numeric.where(mask).compare(df2_numeric.where(mask))
    else:
        result_file.write(
            "The dataframes have different number of rows and columns and can't be compared.\n")
        print("The dataframes have different number of rows and columns and can't be compared.\n")
        return 2, pd.DataFrame()

    # Select only non-numeric columns
    df1_non_numeric = df1.select_dtypes(exclude=[np.number])
    df2_non_numeric = df2.select_dtypes(exclude=[np.number])

    # Compare non-numeric columns
    diff_non_numeric = df1_non_numeric.compare(df2_non_numeric)

    # Concatenate the differences
    diff = pd.concat([diff_numeric, diff_non_numeric])

    # Reset the index and add the old index as a new column
    diff = diff.reset_index().rename(
        columns={'index': 'original_row_number', 'self': 'manual', 'other': 'automated'})

    if not diff.empty:
        result_file.write("The dataframes have differences.\n")
        return 7, diff
    else:
        result_file.write("The dataframes have the same columns and values.\n")
        return 0, diff


def main(_vendor, _filename, _new_file, _ignore_rows, _ignore_data_types, _error_codes):
    """
    Main function to run the script.

    Parameters:
    _vendor (str): The vendor name.
    _filename (str): The filename.
    _new_file (int): Whether to start a new results file or append to an existing one.
    _ignore_rows (int): Whether to ignore the number of rows in the comparison.
    _ignore_data_types (int): Whether to ignore the data types in the comparison.
    _error_codes (dict): The dictionary of error codes.
    """
    # Set the vendor and filename
    # Set the filenames for the manual and automated files
    filenamea = _filename + 'a.xlsx'
    filenamem = _filename + 'm.xlsx'

    # Set the columns to sort the dataframes on
    columns_to_sort = ('Account', 'Product name',
                       'Net billed', 'Gross commission')

    # establish file paths for input and output files
    manual_url = Path(_vendor, "test_files", "rpm_files_manual", filenamem)
    automated_url = Path(_vendor, "test_files",
                         "rpm_files_automation", filenamea)
    output_url = Path(_vendor, "test_files", "output",
                      _filename + '_output.xlsx')
    column_mapping_url = Path(_vendor, "test_files", "column_mapping.xlsx")
    test_results_url = Path(_vendor, "test_files",
                            "output", "test_results.txt")

    # read in the input files
    try:
        manual = pd.read_excel(manual_url)
    except:
        print("Error opening manual file")

    try:
        automated = pd.read_excel(automated_url)
    except:
        print("ERROR: opening automated file")

    try:
        column_mapping = pd.read_excel(column_mapping_url)
    except:
        print("ERROR: opening mapping file")

    try:
        if _new_file:
            test_results = open(test_results_url, 'wt', encoding='utf-8')
            test_results.write(f"Test results for '{_vendor}'\n")
        else:
            test_results = open(test_results_url, 'at', encoding='utf-8')
    except:
        print("ERROR: opening test results file")

    test_results.write(
        f"\n\n{'-'*37}\nTest results for '{_vendor}' - '{_filename}' - {datetime.now()}\n")

    # delete unwanted columns
    manual, automated = basic_conditioning(manual, automated)

    # run basic validation on the dataframes
    error, output_df = basic_validation(
        manual, automated, test_results, _ignore_rows=_ignore_rows, _ignore_data_types=_ignore_data_types)

    match error:
        # If the dataframes do not pass basic validations and the process is not allowed to continue
        case 1 | 2 | 4 | 5:
            print(_error_codes[error])
            print(output_df)
            print("ERROR: Dataframes failed basic validation")
            test_results.write("ERROR: Dataframes failed basic validation")
            test_results.write(f'{_error_codes[error]} \n')
            test_results.write(output_df.to_string())

        # If the dataframes pass basic validations or it is allowed to continue despite error
        case 3 | 6 | 9:
            # Remap columns in manual and automated based on the mapping dataframe
            #  this is done to ensure that all files are sorted and compared in the same way
            manual = rename_and_drop_columns(manual, column_mapping)
            automated = rename_and_drop_columns(automated, column_mapping)

            # Sort columns on both dataframes to compare
            manual_sort_error, manual = sort_rows(manual, columns_to_sort)
            automated_sort_error, automated = sort_rows(
                automated, columns_to_sort)

            # Compare the sorted dataframes
            if manual_sort_error == 0 and automated_sort_error == 0:
                differences_error, df_differences = compare_df(
                    manual, automated, result_file=test_results)
                print(_error_codes[differences_error])
                print(df_differences)
                test_results.write(
                    f'{_error_codes[differences_error]} \n Differences: \n')
                test_results.write(df_differences.to_string())
                # Produce output files
                with pd.ExcelWriter(output_url) as writer:
                    manual.to_excel(writer, sheet_name='manual', index=False)
                    automated.to_excel(
                        writer, sheet_name='automated', index=False)
            elif manual_sort_error == 8 or automated_sort_error == 8:
                print(_error_codes[8])
                test_results.write(_error_codes[8])
            else:
                print('Unknown error while sorting dataframes')
                test_results.write('Unknown error while sorting dataframes')
    test_results.close()


# Run the main function
if __name__ == "__main__":
    # vendor = input ("Vendor: ")
    vendor = 'Sandler'

    # filename = input ("Filename (mmm-yy): ")
    filename = 'feb-24'

    # Set variables
    # Set variable to 1 to start a new results file, or to 0 to append to an existing file
    NEW_FILE = 0
    # Set variable to 1 to ignore the number of rows in the comparison, or to 0 to compare the number of rows
    IGNORE_ROWS = 0
    # Set variable to 1 to ignore the data type in the comparison, or to 0 to compare the data type
    IGNORE_DATA_TYPES = 0

    # Set dictionary of error codes
    error_codes = {
        0: 'The dataframes have the same columns and values.',
        1: 'The dataframes do not have the same number of columns.',
        2: 'The dataframes have the same number of columns, but different number of rows. Process stopped.',
        3: 'The dataframes have the same number of columns, but different number of rows. Ignoring the number of rows in the comparison.',
        4: 'The dataframes have the same number of columns, different columns names.',
        5: 'The dataframe have the same columns, but different data types.',
        6: 'The dataframes have the same columns and data types.',
        7: 'The dataframes have different values.',
        8: 'One or more columns in columns_to_sort do not exist in the dataframe.',
        9: 'The dataframe have the same columns, but different data types. Ignoring the data types in the comparison.',
        10: 'The dataframes have different total sum of Gross commission for each Account.'
    }
    main(_vendor=vendor, _filename=filename, _new_file=NEW_FILE, _ignore_rows=IGNORE_ROWS,
         _ignore_data_types=IGNORE_DATA_TYPES, _error_codes=error_codes)
