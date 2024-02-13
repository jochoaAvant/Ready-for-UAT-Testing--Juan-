import pandas as pd
import numpy as np
from pathlib import Path


def rename_and_drop_columns(input_df, column_mapping_df):
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
    # Remove 'Id' column and 'VIA' Column if any exists
    if any(column_name == 'id' for column_name in df2.columns):
        df2.drop(['id'], axis=1, inplace=True)
    if any(column_name == 'VIA' for column_name in df1.columns):
        df1.drop(['VIA'], axis=1, inplace=True)
    return df1, df2


def basic_validation(df1, df2, result_file, _ignore_rows=0):
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
                return error
        else:
            error = 1
            return error

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
                    f"The column '{column}' has different data types in the dataframes.")
                different_types = 1
                new_row = pd.DataFrame([{'Column': column, 'Same DataType': 'No',
                                         'DataTypeManual': df1[column].dtype,
                                         'DataTypeAutomated': df2[column].dtype}])
                comparison_df = pd.concat([comparison_df,
                                          new_row], ignore_index=True)
        if different_types == 1:
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
        return error, different_columns


def sort_rows(unsorted_df, columns_to_sort):
    # Check if all columns in columns_to_sort exist in the dataframe
    if all(item in unsorted_df.columns for item in columns_to_sort):
        # Sort the dataframe
        df_sorted = unsorted_df.sort_values(
            list(columns_to_sort)).reset_index(drop=True)
        return (0, df_sorted)
    else:
        return (8, df_sorted)


def compare_df(df1, df2, result_file):
    # Select only numeric columns
    df1_numeric = df1.select_dtypes(
        include=[np.number]).reset_index(drop=True)
    df2_numeric = df2.select_dtypes(
        include=[np.number]).reset_index(drop=True)

    # Create a boolean mask for values in df1_numeric and df2_numeric that are not close
    mask = ~np.isclose(df1_numeric.values, df2_numeric.values, atol=0.01)

    # Apply the mask to df1_numeric and df2_numeric, and drop rows that contain only NaN
    diff_numeric = df1_numeric.where(mask).compare(df2_numeric.where(mask))

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


def main(_vendor, _filename, _new_file, _ignore_rows, _error_codes):
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
        if (_new_file):
            test_results = open(test_results_url, 'wt')
            test_results.write(f"Test results for '{_vendor}'\n")
        else:
            test_results = open(test_results_url, 'at')
    except:
        print("ERROR: opening test results file")

    test_results.write(
        f"\n\n{'-'*37}\nTest results for '{_vendor}' - '{_filename}' \n")

    # delete unwanted columns
    manual, automated = basic_conditioning(manual, automated)

    # run basic validation on the dataframes
    error, output_df = basic_validation(
        manual, automated, test_results, _ignore_rows=_ignore_rows)

    match error:
        case 1 | 2 | 4 | 5:
            print(_error_codes[error])
            print(output_df)
            print("ERROR: Dataframes failed basic validation")
            test_results.write("ERROR: Dataframes failed basic validation")
            test_results.write(f'{error_codes[error]} \n')
            test_results.write(output_df.to_string())

        case 3 | 6:
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
                #     if isinstance(df_differences, pd.DataFrame):
                #         df_differences.to_excel(writer, sheet_name='differences', index=False)
                #     else:
                #         pd.DataFrame([df_differences]).to_excel(writer,
                #               sheet_name='differences', index=False)
            elif manual_sort_error == 8:
                print(_error_codes[manual_sort_error])
                test_results.write(_error_codes[manual_sort_error])
            elif automated_sort_error == 8:
                print(_error_codes[automated_sort_error])
                test_results.write(_error_codes[automated_sort_error])
            else:
                print('Unknown error while sorting dataframes')
                test_results.write('Unknown error while sorting dataframes')
    test_results.close()


# Run the main function
if __name__ == "__main__":
    # vendor = input ("Vendor: ")
    vendor = 'Nitel'

    # filename = input ("Filename (mmm-yy): ")
    filename = 'dec-23'

    # Set variables
    # Set variable to 1 to start a new results file, or to 0 to append to an existing file
    new_file = 0
    ignore_rows = 1    # Set variable to 1 to ignore the number of rows in the comparison, or to 0 to compare the number of rows

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
        8: 'One or more columns in columns_to_sort do not exist in the dataframe.'
    }
    main(_vendor=vendor, _filename=filename, _new_file=new_file, _ignore_rows=ignore_rows,
         _error_codes=error_codes)
