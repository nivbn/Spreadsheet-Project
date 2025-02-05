import pandas as pd
import re
from simpleeval import SimpleEval
from typing import List, Tuple, Dict, Set
import yaml
from spreadsheet_functions import SpreadsheetFunctions
from spreadsheet_history import SpreadsheetHistory
import os

# Constants
ALPHABET_LENGTH = 26
ASCII_A = 65


class Spreadsheet:
    """
    Represents a spreadsheet application that allows for storing data,
    evaluating formulas, functions, and make cell range operations.
    """

    def __init__(self):
        """
        Initializes a new instance of the Spreadsheet class.
        """
        # DataFrame for storing cell data
        self._data: pd.DataFrame = pd.DataFrame()

        # SimpleEval instance for safe expression evaluation
        self._evaluator = SimpleEval()

        # Mapping of cells to their formulas
        self._formulas: Dict[str, str] = {}  # key: cell, value: formula string

        # Mapping of cells to their functions
        self._functions: Dict[str, Tuple[str, Tuple[str, str]]] = {}  # key: cell, value: (function_name, arguments)

        # Mapping of cells to their direct dependencies
        self._dependencies: Dict[str, List[str]] = {}  # key: cell, value: list of dependent cells

        # Mapping of cells to cells that depend on them for updates
        self._reverse_dependencies: Dict[str, Set[str]] = {}  # key: cell, value: list of cells depending on it

        # Set history_dir to a new subdirectory named 'history_files' within the current program directory
        self.history_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'history_files')

        # Ensure the history directory exists, create if it does not
        if not os.path.exists(self.history_dir):
            os.makedirs(self.history_dir)

        # Initialize the history manager with the specified directory for storing state files.
        self.history_manager = SpreadsheetHistory(self.history_dir)

    def create_sheet(self, rows: int, columns: int) -> None:
        """
        Creates a new sheet with the specified number of rows and columns.

        Args:
            rows (int): Number of rows.
            columns (int): Number of columns.
        """
        # Generate column labels as strings (e.g., A, B, C,...AA, BB,...)
        column_labels = self.generate_column_labels(columns)

        # Initialize the DataFrame with the specified dimensions and None as initial value.
        self._data = pd.DataFrame(index=range(1, rows + 1), columns=column_labels, dtype=object)

    def display_sheet(self) -> None:
        """
        Displays the current state of the spreadsheet.
        """
        # Convert DataFrame to strings for consistent data type and avoid down-casting issues
        display_df = self._data.astype(str)

        # Replace 'nan' strings (resulting from NA values converted to strings) with empty strings
        display_df = display_df.replace('nan', '')

        print(display_df)

    @staticmethod
    def parse_cell_address(cell: str) -> Tuple[int, str]:
        """
        Parses a cell address into its row number and column label.

        Args:
            cell (str): The cell address (e.g., 'A1').

        Returns:
            Tuple[int, str]: A tuple containing the row number and column label.
        """
        # Check if valid cell address
        match = re.match(r"^([A-Za-z]+)(\d+)$", cell)
        if not match:
            raise ValueError(f"Invalid cell address format: '{cell}'")
        # Separate the column label and row number
        column_label, row_number = match.groups()
        return int(row_number), column_label.upper()

    def parse_cell_range(self, cell_range: Tuple[str, str]) -> Tuple[slice, slice]:
        """
        Parses a cell range into start and end indices suitable for DataFrame slicing.

        This method is used for operations that involve a range of cells, such as summing
        or averaging values across a specified range.

        Args:
            cell_range (Tuple[str, str]): The start and end cell addresses as a tuple.

        Returns:
            Tuple[slice, slice]: A tuple of slices for row and column indices.
        """
        start_row, start_col = self.parse_cell_address(cell_range[0])
        end_row, end_col = self.parse_cell_address(cell_range[1])

        # Convert column labels to positions for slicing
        col_start_pos = self.column_letter_to_index(start_col)
        col_end_pos = self.column_letter_to_index(end_col)

        return slice(start_row - 1, end_row), slice(col_start_pos, col_end_pos + 1)

    @staticmethod
    def convert_indices_to_cell_address(row: int, col: int) -> str:
        """
        Converts row and column indices to a spreadsheet cell address.

        Args:
            row (int): The row index.
            col (int): The column index.

        Returns:
            str: The cell address (e.g., 'A1', 'AA10').
        """
        column_label = ''
        while col >= 0:
            col, remainder = divmod(col, ALPHABET_LENGTH)
            column_label = chr(ASCII_A + remainder) + column_label
            col -= 1
        row_label = row + 1
        return f"{column_label}{row_label}"

    @staticmethod
    def column_letter_to_index(column_letter: str) -> int:
        """
        Converts a column letter (e.g., 'A', 'AB') to a zero-based column index.

        Args:
            column_letter (str): The column letter.

        Returns:
            int: The zero-based column index.
        """
        column_index = 0
        for char in column_letter:
            column_index = column_index * ALPHABET_LENGTH + (ord(char.upper()) - ord('A') + 1)
        return column_index - 1

    @staticmethod
    def index_to_column_letter(index: int) -> str:
        """
        Converts a zero-based column index to a column letter (e.g., 0 -> 'A', 27 -> 'AB').

        Args:
            index (int): The zero-based column index.

        Returns:
            str: The corresponding column letter.
        """
        column_letter = ''
        while index >= 0:
            index, remainder = divmod(index, ALPHABET_LENGTH)
            column_letter = chr(ASCII_A + remainder) + column_letter
            index -= 1
        return column_letter

    @staticmethod
    def generate_column_labels(n: int) -> List[str]:
        """
        Generates column labels for the _spreadsheet.

        Args:
            n (int): Number of column labels to generate.

        Returns:
            List[str]: Generated column labels.
        """
        labels = []
        for i in range(1, n + 1):
            label = ''
            while i > 0:
                i, remainder = divmod(i - 1, ALPHABET_LENGTH)
                label = chr(ASCII_A + remainder) + label
            labels.append(label)
        return labels

    def cell_in_function_range(self, cell: str, cell_range: Tuple[str, str]) -> bool:
        """
        Determines if a given cell is within a specified range used in a function.

        Args:
            cell (str): The cell address to check (e.g., 'B2').
            cell_range (tuple): The range string used in the function (e.g., ('A1', 'C3')).

        Returns:
            bool: True if the cell is within the range, False otherwise.
        """
        # Split the function's argument range into start and end cell addresses
        start_cell, end_cell = cell_range

        # Parse the start and end cell addresses to get their row numbers and column labels
        start_row, start_col_label = self.parse_cell_address(start_cell)
        end_row, end_col_label = self.parse_cell_address(end_cell)

        # Parse the updated cell's address to get its row number and column label
        updated_row, updated_col_label = self.parse_cell_address(cell)

        # Convert the column labels of the start, end, and updated cells to zero-based indices
        start_col_index = self.column_letter_to_index(start_col_label)
        end_col_index = self.column_letter_to_index(end_col_label)
        updated_col_index = self.column_letter_to_index(updated_col_label)

        # Determine if the updated cell falls within the range of the function's argument
        is_within_row_range = (start_row <= updated_row <= end_row)
        is_within_col_range = (start_col_index <= updated_col_index <= end_col_index)

        # The cell is within the function's range if it falls within both the row and column ranges
        return is_within_row_range and is_within_col_range

    def is_valid_range(self, start_cell: str, end_cell: str) -> bool:
        """
        Validates if the provided cell range is valid by ensuring the start cell comes
        before the end cell in the spreadsheet.

        Args:
            start_cell (str): The start cell address of the range.
            end_cell (str): The end cell address of the range.

        Returns:
            bool: True if the range is valid, False otherwise.
        """
        start_row, start_col = self.parse_cell_address(start_cell)
        end_row, end_col = self.parse_cell_address(end_cell)
        return start_row <= end_row and start_col <= end_col

    def get_cell_value(self, cell: str) -> str:
        """
        Retrieves the value of a specified cell.

        Args:
            cell (str): Cell address.

        Returns:
            str: The value of the cell.
        """
        # Implement logic to translate cell notation to DataFrame indices
        row_label, column_label = self.parse_cell_address(cell)
        try:
            return self._data.at[row_label, column_label]

        except KeyError:
            return "#REF!"  # return "#REF!" to indicate a reference error

    def get_cell_formula(self, cell: str) -> str:
        """
        Retrieves the formula of a specified cell, if it exists.

        Args:
            cell (str): Cell address in the format 'A1', 'B2', etc.

        Returns:
            str: The formula of the cell if it exists, else None.
        """
        # Return the formula if it exists, else return None
        return self._formulas.get(cell, None)

    def update_dependencies(self, cell: str, formula: str) -> None:
        """
        Updates the dependencies of a given cell based on its formula.

        Args:
            cell (str): The cell being updated.
            formula (str): The formula entered into the cell.
        """
        # Find all cell references in the formula
        self._dependencies[cell] = re.findall(r'[A-Za-z]+\d+', formula)
        for dep in self._dependencies[cell]:
            # Update reverse dependencies for efficient recalculation
            if dep in self._reverse_dependencies:
                self._reverse_dependencies[dep].add(cell)
            else:
                self._reverse_dependencies[dep] = {cell}

    def recalculate_dependents(self, cell: str) -> None:
        """
        Recalculates the values of cells dependent on the given cell.

        Args:
            cell (str): The cell that has been updated.
        """
        # Check if any cells depend on this cell
        if cell in self._reverse_dependencies:
            # Recalculate each dependent cell
            for dependent_cell in self._reverse_dependencies[cell]:
                if dependent_cell in self._formulas:
                    # Recalculate the formula of the dependent cell
                    self.calculate_formula(dependent_cell, self._formulas[dependent_cell])
                    # Recursively recalculate any cells dependent on this cell
                    self.recalculate_dependents(dependent_cell)  # Handle cascading dependencies

        # To avoid RuntimeError due to dictionary modification during iteration,
        # iterate over a static list of the dictionary items
        for func_cell, (func_name, cell_range_tuple) in list(self._functions.items()):
            if self.cell_in_function_range(cell, cell_range_tuple):
                # Depending on the function, execute it again to update its value
                if func_name in ["Sum", "Average", "Max", "Min", "Count", "Median", "Product"]:
                    self.execute_function(func_cell, func_name, cell_range_tuple)
                # Ensure we also check for cells dependent on this function cell
                self.recalculate_dependents(func_cell)

    def has_circular_dependency(self, cell: str, target: str, visited: set = None) -> bool:
        """
        Checks for circular dependencies involving the given cell.

        Args:
            cell (str): The cell to check for circular dependencies.
            target (str): The original target cell being updated.
            visited (set, optional): Set of already visited cells in the dependency chain.

        Returns:
            bool: True if a circular dependency is detected; False otherwise.
        """
        if visited is None:
            visited = set()
        visited.add(cell)

        # Check if the cell directly depends on itself
        if cell == target:
            return True

        if cell in self._dependencies:
            for dep in self._dependencies[cell]:
                if dep == target:
                    return True  # Direct circular dependency detected
                if dep not in visited:
                    if self.has_circular_dependency(dep, target, visited):
                        return True
        return False

    def track_function_dependencies(self, cell: str, cell_range: Tuple[str, str]) -> None:
        """
        Tracks the dependencies for a given cell based on its associated function's range.

        Args:
            cell (str): The cell whose dependencies are being updated (e.g., 'A1').
            cell_range (Tuple[str, str]): A tuple containing the start and end cells that define the range
                                          of the function (e.g., ('A2', 'B4')).
        """
        # Parse the start and end cell addresses to get their row and column indices
        start_cell, end_cell = cell_range
        start_row, start_col = self.parse_cell_address(start_cell)
        end_row, end_col = self.parse_cell_address(end_cell)

        # Iterate over each cell in the specified range
        for row in range(start_row, end_row + 1):
            for col in range(self.column_letter_to_index(start_col), self.column_letter_to_index(end_col) + 1):
                # Convert the current row and column indices back to a cell address (e.g., 'C3')
                dep_cell = self.convert_indices_to_cell_address(row - 1, col)

                # Update reverse dependencies to include this cell
                if dep_cell in self._reverse_dependencies:
                    self._reverse_dependencies[dep_cell].add(cell)
                else:
                    self._reverse_dependencies[dep_cell] = {cell}

                # Ensure the main cell's dependencies are initialized as a set before adding
                if cell in self._dependencies:
                    self._dependencies[cell].add(dep_cell)
                else:
                    self._dependencies[cell] = {dep_cell}

    def expand_sheet_to_include_cell(self, cell_address: str) -> None:
        """
        Expands the sheet to accommodate a specified cell address, ensuring that new cells
        are added in a manner consistent with the DataFrame's default behavior,
        and ensuring that no intermediate columns or rows are skipped.

        Args:
            cell_address (str): The cell address (e.g., 'A1', 'B2') of the target cell.
        """
        row_number, column_label = self.parse_cell_address(cell_address)

        # Calculate the target column index without reversing the label
        target_col_index = 0
        for char in column_label:
            target_col_index = target_col_index * ALPHABET_LENGTH + (ord(char) - ord('A') + 1)

        # Check if attempting to exceed the maximum limits
        if target_col_index > 500 or row_number > 500:
            raise ValueError("Exceeding maximum sheet size of 500x500.")

        # Check if the target column and row exist, expand the sheet only if necessary
        if target_col_index > len(self._data.columns) or row_number > len(self._data.index):
            self.expand_columns(target_col_index)
            self.expand_rows(row_number)

        self._data.sort_index(inplace=True)

    def expand_columns(self, target_col_index: int) -> None:
        """Expand the sheet to include up to the target column index."""
        current_max_col = len(self._data.columns)
        if target_col_index > current_max_col:
            # Generate the additional columns needed
            additional_columns = self.generate_column_labels(target_col_index)[
                                 -1 * (target_col_index - current_max_col):]
            self._data = self._data.reindex(columns=self._data.columns.tolist() + additional_columns, fill_value="")

    def expand_rows(self, target_row_number: int) -> None:
        """Expand the sheet to include up to the target row number."""
        if target_row_number > len(self._data.index):
            self._data = self._data.reindex(index=range(1, target_row_number + 1), fill_value="")

    def set_cell_value(self, cell: str, value: str, function_info=None) -> None:
        """
        Sets the value of a specified cell and triggers updates for dependent cells.

        This method handles both direct value assignments and formula evaluations. For formulas,
        it checks for circular dependencies to prevent infinite loops. If a formula is specified,
        the method evaluates it, updates the cell value, and recalculates any dependent cells.

        Args:
            cell (str): The address of the cell to update (e.g., 'A1').
            value (str): The new value or formula for the cell. Formulas must start with '='.
            function_info: The function used to update the cell value (if been used).
        """
        # Ensure the spreadsheet includes the cell for both direct values and formulas
        self.expand_sheet_to_include_cell(cell)

        # If value is an empty string, clear the cell's value, formula, and function info
        if value == "":
            self.delete_cell(cell)
        # Clear any existing formula or function if a new value is being set.
        if cell in self._formulas or cell in self._functions:
            self.delete_cell(cell)
        # If function_info is provided, store it. Example: ("Sum", ("A1","A9"))
        if function_info:
            self._functions[cell] = function_info

        # Ensure that the value is a string
        value_str = str(value)
        try:
            # Attempt to convert value to a float for numerical operations.
            if not value_str.startswith("="):
                true_value = float(value_str)
            else:
                true_value = value_str  # Keep as formula if conversion is not applicable
        except ValueError:
            true_value = value_str  # Keep as string if conversion fails

        # Check if the value is a formula (starts with "=")
        if value_str.startswith("="):
            # Handle formula evaluation
            formula = value_str[1:]  # Remove '=' sign.
            temp_dependencies = re.findall(r'[A-Za-z]+\d+', formula)

            # Check for circular dependency before updating
            for dep in temp_dependencies:
                if self.has_circular_dependency(dep, cell):
                    raise Exception(f"Circular dependency detected involving {cell} and {dep}.")

            # update the cell's formula and dependencies
            self._formulas[cell] = formula
            self.update_dependencies(cell, formula)
            # Evaluate the formula and update the cell value accordingly
            self.calculate_formula(cell, formula)

        else:
            # Directly set the cell value for non-formula values
            self.set_cell_value_direct(cell, true_value)

        self.recalculate_dependents(cell)  # Trigger recalculation of dependent cells, if any

    def set_cell_value_direct(self, cell: str, value):
        """
        Directly sets the value of a cell without evaluating formulas or dependencies.

        Args:
            cell (str): The cell to update.
            value: The value to set in the cell.
        """
        row_label, column_label = self.parse_cell_address(cell)

        # If value is a float and its decimal part is 0, convert it to an integer
        if isinstance(value, float) and value.is_integer():
            value = int(value)

        self._data.at[row_label, column_label] = value

    def delete_cell(self, cell: str) -> None:
        """
        Deletes the value and any formula or function information of a specified cell,
        and triggers updates for any cells that depend on this cell.

        Args:
            cell (str): The address of the cell to delete (e.g., 'A1').
        """
        # Reset the cell's value to its default
        self.set_cell_value_direct(cell, None)

        # If the cell has formulas or functions, remove them since the cell is being 'cleared'
        self._formulas.pop(cell, None)
        self._functions.pop(cell, None)

        # Trigger recalculation of dependent cells as their dependency
        self.recalculate_dependents(cell)

    def enter_data(self, cells_range: Tuple[str, str], value: str) -> None:
        """
            Enters data or formulas into a specified range of cells.

            This method allows for batch updating of cell values or applying formulas
            within a specified range, useful for initializing or modifying large areas
            of the sheet at once.

            Args:
                cells_range (Tuple[str, str]): The start and end addresses of the cell range.
                value (str): The value or formula to be entered into the cells.
            """
        # Parsing the cell range to get start and end indices for rows and columns
        start_row, start_col = self.parse_cell_address(cells_range[0])
        end_row, end_col = self.parse_cell_address(cells_range[1])

        # Column labels are converted to indices for iteration
        start_col_index = self.column_letter_to_index(start_col)
        end_col_index = self.column_letter_to_index(end_col)

        # Check if the range is valid
        if not self.is_valid_range(cells_range[0], cells_range[1]):
            raise ValueError(f"Invalid cell range: '{cells_range[0]}' is after '{cells_range[1]}'.")

        # Iterate over each cell in the specified range
        for row in range(start_row, end_row + 1):
            for col_index in range(start_col_index, end_col_index + 1):
                # Convert back to column label
                col_label = self.index_to_column_letter(col_index)
                cell_address = f"{col_label}{row}"
                self.set_cell_value(cell_address, value)

    def resolve_cell_references(self, formula: str) -> None:
        """
        Resolves all cell references within a formula to their current values.

        This method updates the SimpleEval evaluator's names dictionary,
        mapping each cell reference found in the formula to its respective value.
        This allows the formula to be evaluated with the current cell values.

        Args:
            formula (str): The formula containing cell references to resolve.
        """
        # Find all cell references in the formula using "re" library.
        cell_refs = re.findall(r'[A-Za-z]+\d+', formula)

        for ref in cell_refs:
            # Get the current value of the cell reference.
            cell_value = self.get_cell_value(ref) or 0  # Default to 0 if the cell is empty or not found.

            # Update the evaluator's context with the resolved cell value.
            self._evaluator.names[ref] = cell_value

    def calculate_formula(self, target_cell: str, formula: str) -> None:
        """
        Evaluates a formula and updates the specified target cell with the result.

        Before evaluation, all cell references within the formula are resolved to
        their current values. The formula is then evaluated, and the target cell
        is updated with the calculated result.

        Args:
            target_cell (str): The address of the cell to update with the formula result.
            formula (str): The formula to evaluate.
        """
        # Resolve cell references in the formula to current values.
        self.resolve_cell_references(formula)

        try:
            # Evaluate the formula with resolved cell references, using safely eval method from "simpleeval" library .
            result = self._evaluator.eval(formula)
            # Update the target cell with the result
            self.set_cell_value_direct(target_cell, result)

        except ZeroDivisionError:
            # Handle division by zero error specifically
            result = "div/0 Error"
            # Update the target cell with an error message
            self.set_cell_value_direct(target_cell, result)

        except Exception as e:

            # Handle value errors, operations on incompatible types.
            if 'unsupported operand type(s)' in str(e) or "can't multiply sequence by non-int" in str(e):
                result = "Value Error"
            else:
                result = "General Error"
            # Update the target cell with an error message
            self.set_cell_value_direct(target_cell, result)
            # Print an error message if the formula evaluation fails.
            raise Exception(f"Error evaluating formula '{formula}': {e}")

    def function_handle(self, cell: str, function_name: str, cell_range: Tuple[str, str], result_value: str) -> None:
        """
        Handles the executing of a spreadsheet function by setting the cell's value
        to the function's result and tracking the dependencies for the function.

        Args:
            cell (str): The cell address where the function's result is displayed (e.g., 'A1').
            function_name (str): The name of the function being executed (e.g., 'Sum').
            cell_range (Tuple[str, str]): The range of cells the function operates on (e.g., ('A2', 'A5')).
            result_value (str): The result of the function calculation as a string.
        """
        # Package the function name and range into a tuple for storing as function information
        function_info = (function_name, cell_range)
        # Set the cell's value to the function's result and store the function information
        self.set_cell_value(cell, result_value, function_info=function_info)
        # Track the dependencies this function has on other cells in the range
        self.track_function_dependencies(cell, cell_range)

    def execute_function(self, cell: str, function_name: str, cell_range: Tuple[str, str]) -> None:
        """
        Executes a specified function over a given cell range and updates the target cell
        with the result of the function.

        Args:
            cell (str): The target cell address where the function result is to be displayed.
            function_name (str): The name of the function to execute (e.g., 'Sum', 'Average').
            cell_range (Tuple[str, str]): The start and end cell addresses defining the range
                                          over which the function is to be applied.
        """
        match = re.match(r"^([A-Za-z]+)(\d+)$", cell)
        if not match:
            raise ValueError(f"Invalid cell address format: '{cell}'")

        # Check if both start and end cell addresses are provided
        if not cell_range or len(cell_range) != 2 or not all(cell_range):
            raise ValueError("Both start and end cell addresses must be provided WHen using this function.")

        # Verify that the start cell comes before the end cell
        start_cell, end_cell = cell_range
        if not self.is_valid_range(start_cell, end_cell):
            raise ValueError(f"Invalid cell range: '{start_cell}' is after '{end_cell}'.")

        # Check if the target cell is in the start and end cell range
        if self.cell_in_function_range(cell, cell_range):
            raise ValueError(f"Target cell address '{cell}' is in function range '{cell_range}'.")

        # Determine the function to execute based on its name and calculate the result
        if function_name == "Sum":
            result = SpreadsheetFunctions.function_sum(cell_range, self)
        elif function_name == "Average":
            result = SpreadsheetFunctions.function_average(cell_range, self)
        elif function_name == "Max":
            result = SpreadsheetFunctions.function_max(cell_range, self)
        elif function_name == "Min":
            result = SpreadsheetFunctions.function_min(cell_range, self)
        elif function_name == "Count":
            result = SpreadsheetFunctions.function_count(cell_range, self)
        elif function_name == "Median":
            result = SpreadsheetFunctions.function_median(cell_range, self)
        elif function_name == "Product":
            result = SpreadsheetFunctions.function_product(cell_range, self)
        else:
            raise ValueError(f"Unsupported function: {function_name}")

        # Handle the function's execution details such as setting the result and tracking dependencies
        self.function_handle(cell, function_name, cell_range, result)

    def clear_sheet(self) -> None:
        """
        Clears the current sheet, resetting all values to None but preserving the number of rows and columns.
        """
        # Get the current number of rows and columns
        current_rows, current_columns = self._data.shape

        # Call create_sheet with the current dimensions to reset the sheet
        # while preserving its structure
        self.create_sheet(current_rows, current_columns)

        # reset formulas, dependencies, and reverse dependencies
        self._formulas = {}
        self._functions = {}
        self._dependencies = {}
        self._reverse_dependencies = {}

    def save_sheet(self, filename: str):
        """
        Save the current sheet to a file in a YAML format, including formulas and dependencies.
        Args:
            filename (str): The path to the file where the sheet will be saved.
        """
        data_dict = {
            'data': self._data.values.tolist(),
            'columns': self._data.columns.tolist(),
            'formulas': self._formulas,
            'functions': {k: [v[0], list(v[1])] for k, v in self._functions.items()},  # Convert tuple to list
            'dependencies': self._dependencies,
            'reverse_dependencies': self._reverse_dependencies,
        }

        with open(filename, 'w') as file:
            yaml.dump(data_dict, file, default_flow_style=False)

    def load_sheet(self, filename: str):
        """
        Load a sheet from a YAML file, including formulas, functions, and dependencies.
        Args:
            filename (str): The path to the file from which the sheet will be loaded.
        """
        with open(filename, 'r') as file:
            data_dict = yaml.safe_load(file)

        # Check for the main data part
        if data_dict and 'data' in data_dict and 'columns' in data_dict:
            self._data = pd.DataFrame(data=data_dict['data'], columns=data_dict['columns'])
            self._data.index = pd.RangeIndex(start=1, stop=len(self._data) + 1, step=1)
        else:
            raise Exception("Loading Error, The loaded file does not contain the expected data structure.")

        # Load formulas
        self._formulas = data_dict.get('formulas', {})

        # Load functions, converting lists back to tuples if necessary
        functions = data_dict.get('functions', {})
        self._functions = {k: (v[0], tuple(v[1])) if isinstance(v, list) and len(v) > 1 else v for k, v in
                           functions.items()}

        # Load dependencies and reverse dependencies.
        self._dependencies = data_dict.get('dependencies', {})
        self._reverse_dependencies = data_dict.get('reverse_dependencies', {})
