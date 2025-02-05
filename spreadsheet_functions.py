import pandas as pd
from typing import Tuple


class SpreadsheetFunctions:
    """
    Implements common spreadsheet functions, extending the capabilities of the Spreadsheet class.
    """

    @staticmethod
    def function_count(cell_range: Tuple[str, str], spreadsheet) -> str:
        """
        Counts the number of cells in a specified range that contain numeric values.

        Args:
            cell_range (Tuple[str, str]): The start and end addresses of the cell range.
            spreadsheet: The Spreadsheet instance to operate on.

        Returns:
            int: The count of numeric values within the specified range.
        """
        row_slice, col_slice = spreadsheet.parse_cell_range(cell_range)
        numeric_cells = spreadsheet._data.iloc[row_slice, col_slice].apply(pd.to_numeric, errors='coerce').notnull()
        count_value = numeric_cells.sum().sum()
        return str(count_value)

    @staticmethod
    def function_sum(cell_range: Tuple[str, str], spreadsheet) -> str:
        """
        Calculates the sum of numeric values in the specified cell range.

        Args:
            cell_range (Tuple[str, str]): The start and end addresses defining the cell range.
            spreadsheet: The Spreadsheet instance to operate on.

        Returns:
            str: Sum of numeric values within the specified range as a string.
        """
        row_slice, col_slice = spreadsheet.parse_cell_range(cell_range)
        sum_value = spreadsheet._data.iloc[row_slice, col_slice].apply(pd.to_numeric, errors='coerce').sum().sum()
        return str(sum_value)

    @staticmethod
    def function_average(cell_range: Tuple[str, str], spreadsheet) -> str:
        """
        Calculates the average of numeric values in the specified cell range.

        Args:
            cell_range (Tuple[str, str]): The start and end addresses defining the cell range.
            spreadsheet: The Spreadsheet instance to operate on.

        Returns:
            str: Average of numeric values within the specified range as a string.
        """
        row_slice, col_slice = spreadsheet.parse_cell_range(cell_range)
        avg_value = spreadsheet._data.iloc[row_slice, col_slice].apply(pd.to_numeric, errors='coerce').mean().mean()
        return str(avg_value)

    @staticmethod
    def function_max(cell_range: Tuple[str, str], spreadsheet) -> str:
        """
        Finds the maximum numeric value in the specified cell range.

        Args:
            cell_range (Tuple[str, str]): The start and end addresses defining the cell range.
            spreadsheet: The Spreadsheet instance to operate on.

        Returns:
            str: Maximum numeric value within the specified range as a string.
        """
        row_slice, col_slice = spreadsheet.parse_cell_range(cell_range)
        max_value = spreadsheet._data.iloc[row_slice, col_slice].apply(pd.to_numeric, errors='coerce').max().max()
        return str(max_value)

    @staticmethod
    def function_min(cell_range: Tuple[str, str], spreadsheet) -> str:
        """
        Finds the minimum numeric value in the specified cell range.

        Args:
            cell_range (Tuple[str, str]): The start and end addresses defining the cell range.
            spreadsheet: The Spreadsheet instance to operate on.

        Returns:
            str: Minimum numeric value within the specified range as a string.
        """
        row_slice, col_slice = spreadsheet.parse_cell_range(cell_range)
        min_value = spreadsheet._data.iloc[row_slice, col_slice].apply(pd.to_numeric, errors='coerce').min().min()
        return str(min_value)

    @staticmethod
    def function_median(cell_range: Tuple[str, str], spreadsheet) -> str:
        """
        Computes the median of the numeric values in the specified cell range.

        Args:
            cell_range (Tuple[str, str]): The start and end addresses defining the cell range.
            spreadsheet: The Spreadsheet instance to operate on.

        Returns:
            str: Median of numeric values within the specified range as a string.
        """
        row_slice, col_slice = spreadsheet.parse_cell_range(cell_range)
        median_value = (spreadsheet._data.iloc[row_slice, col_slice].apply
                        (pd.to_numeric, errors='coerce').median().median())
        return str(median_value)

    @staticmethod
    def function_product(cell_range: Tuple[str, str], spreadsheet) -> str:
        """
        Computes the product of the numeric values in the specified cell range.

        Args:
            cell_range (Tuple[str, str]): The start and end addresses defining the cell range.
            spreadsheet: The Spreadsheet instance to operate on.

        Returns:
            str: Product of numeric values within the specified range as a string.
        """
        row_slice, col_slice = spreadsheet.parse_cell_range(cell_range)
        product_value = spreadsheet._data.iloc[row_slice, col_slice].apply(pd.to_numeric, errors='coerce').prod().prod()
        return str(product_value)
