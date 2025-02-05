from spreadsheet import *
import unittest


class TestSpreadsheet(unittest.TestCase):
    """
     Tests for the functionality of the Spreadsheet class. These tests cover a range of functionalities,
     including basic cell operations, formula evaluation, handling circular dependencies, setting and getting formulas,
     handling out-of-bounds references, operations with non-numeric values, complex formula evaluations, error handling
     in formulas, resetting cell values, dependencies and cascading updates when values change, handling large numbers
     and precision, deleting cells and observing the impact on dependents, and the implementation of various spreadsheet
     functions like Sum, Average, Max, Min, Count, Median, and Product. Additionally, tests for function execution with
     invalid inputs, ranges, and addressing are included to ensure robust error handling.
     """

    def setUp(self):
        self.sheet = Spreadsheet()

    def test_basic_cell_operations(self):
        self.sheet.create_sheet(2, 2)

        self.sheet.set_cell_value("A1", "100")
        assert self.sheet.get_cell_value("A1") == 100.0, "A1 should have a value of 100.0"

        self.sheet.set_cell_value("B2", "200")
        assert self.sheet.get_cell_value("B2") == 200.0, "B2 should have a value of 200.0"

    def test_get_cell_formula_not_set(self):
        self.sheet.create_sheet(2, 2)
        formula = self.sheet.get_cell_formula("A1")
        self.assertIsNone(formula, "Expected no formula to be returned for cell A1")

    def test_formula_evaluation(self):
        self.sheet.create_sheet(3, 1)

        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.set_cell_value("A3", "=A1+A2")
        assert self.sheet.get_cell_value("A3") == 30.0, "A3 should evaluate to 30.0"

    def test_circular_dependency(self):
        self.sheet.create_sheet(2, 1)

        self.sheet.set_cell_value("A1", "=A2*2")
        # Use assertRaises to catch the expected exception
        with self.assertRaises(Exception) as context:
            self.sheet.set_cell_value("A2", "=A1*2")

        # Verify that the exception message matches the expected message
        self.assertEqual(str(context.exception), "Circular dependency detected involving A2 and A1.")

    def test_setting_value_to_previously_formula_cell(self):
        """Test that setting a direct value to a cell that had a formula removes the formula."""
        self.sheet.create_sheet(3, 3)

        # Set a formula and then set a direct value
        self.sheet.set_cell_value("A1", "=A2+A3")
        self.sheet.set_cell_value("A1", "100")

        # Check that the formula is removed and the cell contains the direct value
        self.assertEqual(self.sheet.get_cell_value("A1"), 100.0, "A1 should contain the direct value")
        self.assertIsNone(self.sheet._formulas.get("A1"), "A1's formula should be removed")

    def test_out_of_bounds_reference(self):
        self.sheet.create_sheet(1, 1)

        self.sheet.set_cell_value("A1", "=B2")
        assert self.sheet.get_cell_value("A1") in [0, "#REF!"], "A1 should handle out-of-bounds reference gracefully"

    def test_non_numeric_value_operations(self):
        self.sheet.create_sheet(2, 2)

        self.sheet.set_cell_value("A1", "Hello")
        assert self.sheet.get_cell_value("A1") == "Hello", "A1 should store and return non-numeric values correctly"

        self.sheet.set_cell_value("B2", "World")
        assert self.sheet.get_cell_value("B2") == "World", "B2 should store and return non-numeric values correctly"

    def test_complex_formulas(self):
        self.sheet.create_sheet(3, 3)

        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("B1", "5")
        self.sheet.set_cell_value("C1", "=A1*(B1+5)/2")

        assert self.sheet.get_cell_value("C1") == 50.0, "C1 should correctly calculate the complex formula"

    def test_formula_with_error_value(self):
        self.sheet.create_sheet(2, 2)

        self.sheet.set_cell_value("A1", "0")
        self.sheet.set_cell_value("B1", "=10/A1")  # This should result in a div/0 Error

        # Assuming your implementation uses a specific error message for division by zero
        assert self.sheet.get_cell_value("B1") == "div/0 Error", "B1 should return a div/0 Error"

    def test_cell_value_reset(self):
        self.sheet.create_sheet(2, 2)

        self.sheet.set_cell_value("A1", "100")
        assert self.sheet.get_cell_value("A1") == 100.0, "A1 should initially have a value of 100"

        # Reset A1 to a new value
        self.sheet.set_cell_value("A1", "200")
        assert self.sheet.get_cell_value("A1") == 200.0, "A1 should be updated to a new value of 200"

    def test_updating_cell_affects_dependent_formula(self):
        self.sheet.create_sheet(3, 1)

        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.set_cell_value("A3", "=A1+A2")

        assert self.sheet.get_cell_value("A3") == 30.0, "A3 initial value should be 30"

        # Update A1 and check if A3 updates correctly
        self.sheet.set_cell_value("A1", "15")
        assert self.sheet.get_cell_value("A3") == 35.0, "A3 should update to 35 after A1 is changed"

    def test_multiple_dependencies(self):
        self.sheet.create_sheet(3, 3)
        self.sheet.set_cell_value("A1", "2")
        self.sheet.set_cell_value("B1", "3")
        self.sheet.set_cell_value("C1", "=A1+B1")

        # Initial state assertion - Expecting a float for numerical calculations
        assert self.sheet.get_cell_value("C1") == 5.0, "C1 initial calculation is incorrect"

        self.sheet.set_cell_value("A1", "5")  # Change one of the dependencies

        # After changing A1 - Expecting updated calculation
        assert self.sheet.get_cell_value("C1") == 8.0, "C1 should update to 8.0 after changing A1"

        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.set_cell_value("A3", "=A1+A2")

        # Checking value after additional updates
        assert self.sheet.get_cell_value("A3") == 30.0, "A3 initial value should be 30.0"

        # Update A1 and check if A3 updates correctly
        self.sheet.set_cell_value("A1", "15")
        assert self.sheet.get_cell_value("A3") == 35.0, "A3 should update to 35.0 after A1 is changed"

    def test_cascading_updates(self):
        self.sheet.create_sheet(4, 4)
        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "=A1*2")
        self.sheet.set_cell_value("A3", "=A2+10")
        self.sheet.set_cell_value("A4", "=A3/2")

        # Initial state assertion
        assert self.sheet.get_cell_value("A4") == 15.0, "A4 initial calculation is incorrect"

        self.sheet.set_cell_value("A1", "20")  # This change should cascade

        # After changing A1
        assert self.sheet.get_cell_value("A4") == 25.0, "A4 should update to 25.0 after changing A1"

    def test_complex_formulas_dependency(self):
        self.sheet.create_sheet(5, 5)
        self.sheet.set_cell_value("A1", "100")
        self.sheet.set_cell_value("B1", "200")
        self.sheet.set_cell_value("C1", "300")
        self.sheet.set_cell_value("D1", "=A1+B1*C1-100")
        self.sheet.set_cell_value("E1", "=D1/A1")

        # Initial state assertion
        # Adjust these based on the expected format (string vs. numeric) of your implementation
        assert self.sheet.get_cell_value("D1") == 60000.0, "D1 initial calculation is incorrect"
        assert self.sheet.get_cell_value("E1") == 600.0, "E1 initial calculation is incorrect"

        self.sheet.set_cell_value("A1", "50")  # Update A1 to see how D1 and E1 change

        # After changing A1
        assert self.sheet.get_cell_value("E1") == 1199.0, "E1 should update correctly after changing A1"

    def test_large_numbers_and_precision(self):
        self.sheet.create_sheet(1, 2)

        large_number = 12345678901234567890
        self.sheet.set_cell_value("A1", str(large_number))
        self.assertEqual(self.sheet.get_cell_value("A1"), float(large_number), "A1 should correctly handle large numbers")

        precision_test = "0.12345678901234567890"
        self.sheet.set_cell_value("B1", precision_test)
        self.assertAlmostEqual(self.sheet.get_cell_value("B1"), float(precision_test), places=10,
                               msg="B1 should maintain precision")

    def test_deleting_cells_affects_dependents(self):
        self.sheet.create_sheet(3, 1)

        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "=A1*2")
        self.sheet.delete_cell("A1")
        self.assertTrue(self.sheet.get_cell_value("A2") in ["#REF!", 0],
                        "A2 should reflect an error or reset after A1 is deleted")

    def test_sum_function(self):
        self.sheet.create_sheet(2, 1)
        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.execute_function("A3", "Sum", ("A1", "A2"))
        self.assertEqual(self.sheet.get_cell_value("A3"), 30, "Sum function failed")

    def test_average_function(self):
        self.sheet.create_sheet(2, 1)
        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.execute_function("A3", "Average", ("A1", "A2"))
        self.assertEqual(self.sheet.get_cell_value("A3"), 15, "Average function failed")

    def test_max_function(self):
        self.sheet.create_sheet(2, 1)
        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.execute_function("A3", "Max", ("A1", "A2"))
        self.assertEqual(self.sheet.get_cell_value("A3"), 20, "Max function failed")

    def test_min_function(self):
        self.sheet.create_sheet(2, 1)
        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.execute_function("A3", "Min", ("A1", "A2"))
        self.assertEqual(self.sheet.get_cell_value("A3"), 10, "Min function failed")

    def test_count_function(self):
        self.sheet.create_sheet(3, 1)
        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.set_cell_value("A3", "")
        self.sheet.execute_function("B1", "Count", ("A1", "A3"))
        self.assertEqual(self.sheet.get_cell_value("B1"), 2, "Count function failed")

    def test_median_function(self):
        self.sheet.create_sheet(3, 1)
        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("A2", "20")
        self.sheet.set_cell_value("A3", "30")
        self.sheet.execute_function("B1", "Median", ("A1", "A3"))
        self.assertEqual(self.sheet.get_cell_value("B1"), 20, "Median function failed")

    def test_product_function(self):
        self.sheet.create_sheet(3, 1)
        self.sheet.set_cell_value("A1", "2")
        self.sheet.set_cell_value("A2", "3")
        self.sheet.set_cell_value("A3", "4")
        self.sheet.execute_function("B1", "Product", ("A1", "A3"))
        self.assertEqual(self.sheet.get_cell_value("B1"), 24, "Product function failed")

    def test_unsupported_function(self):
        self.sheet.create_sheet(2, 2)
        with self.assertRaises(ValueError) as context:
            self.sheet.execute_function("B1", "NonExistentFunction", ("A1", "A2"))
        self.assertTrue("Unsupported function" in str(context.exception),
                        "Attempting to execute an unsupported function should raise a ValueError")

    def test_invalid_range_function(self):
        self.sheet.create_sheet(2, 2)
        with self.assertRaises(ValueError) as context:
            self.sheet.execute_function("B1", "Sum", ("A2", "A1"))
        self.assertTrue("Invalid cell range" in str(context.exception),
                        "Providing an invalid cell range should raise a ValueError")

    def test_execute_function_with_invalid_cell_address(self):
        """Test executing a function with an invalid cell address format."""
        self.sheet.create_sheet(3, 3)
        with self.assertRaises(ValueError):
            self.sheet.execute_function("InvalidCellAddress", "Sum", ("A1", "A2"))

    def test_execute_function_without_both_cell_addresses(self):
        """Test executing a function without both start and end cell addresses provided."""
        self.sheet.create_sheet(3, 3)
        with self.assertRaises(ValueError):
            self.sheet.execute_function("A3", "Sum", ("A1",))

    def test_execute_function_with_invalid_range(self):
        """Test executing a function with the start cell coming after the end cell."""
        self.sheet.create_sheet(3, 3)
        with self.assertRaises(ValueError):
            self.sheet.execute_function("A3", "Sum", ("A2", "A1"))

    def test_execute_function_with_target_cell_in_range(self):
        """Test executing a function where the target cell is within the function's range."""
        self.sheet.create_sheet(3, 3)
        with self.assertRaises(ValueError):
            self.sheet.execute_function("A2", "Sum", ("A1", "A3"))

    def test_valid_function_execution(self):
        """Test valid execution of a function to ensure not all cases raise errors."""
        self.sheet.create_sheet(3, 3)
        try:
            self.sheet.execute_function("A4", "Sum", ("A1", "A3"))
        except ValueError:
            self.fail("execute_function raised ValueError unexpectedly!")

    def test_track_function_dependencies(self):
        self.sheet.create_sheet(5, 5)  # Create a 5x5 sheet

        # Define a function range and the cell that depends on this range
        function_cell = 'A1'
        function_range = ('B2', 'C3')

        # Manually set dependencies for the function_cell
        self.sheet._dependencies[function_cell] = set()
        # Track dependencies based on a function's range
        self.sheet.track_function_dependencies(function_cell, function_range)

        # Expected dependencies for function_cell ('A1') are cells within the range 'B2:C3'
        expected_dependencies = {'B2', 'B3', 'C2', 'C3'}
        self.assertEqual(self.sheet._dependencies[function_cell], expected_dependencies,
                         "Function cell dependencies not tracked correctly")

        # Check reverse dependencies for each cell in the range
        for dep_cell in expected_dependencies:
            self.assertIn(function_cell, self.sheet._reverse_dependencies[dep_cell],
                          f"{dep_cell} does not list {function_cell} as a reverse dependency")

        # Additionally, test behavior when adding another dependent cell to ensure dependencies are appended,
        # not overwritten
        additional_function_cell = 'A2'
        self.sheet._dependencies[additional_function_cell] = set()
        self.sheet.track_function_dependencies(additional_function_cell, function_range)

        # Now, cells in the range should have both 'A1' and 'A2' as reverse dependencies
        for dep_cell in expected_dependencies:
            self.assertTrue({function_cell, additional_function_cell}.issubset(self.sheet._reverse_dependencies[dep_cell]),
                            f"{dep_cell} reverse dependencies do not include both {function_cell} and"
                            f" {additional_function_cell}")

    def test_expand_sheet_to_include_cell(self):
        self.sheet.create_sheet(3, 3)

        self.sheet.expand_sheet_to_include_cell('D4')
        self.assertIn('D', self.sheet._data.columns, "Sheet did not expand to include column 'D'")
        self.assertEqual(len(self.sheet._data.index), 4, "Sheet did not expand to include row 4")

        # Verify that the cell at 'D4' can be accessed without error
        try:
            self.sheet.set_cell_value('D4', '100')
            cell_value = self.sheet.get_cell_value('D4')
            self.assertEqual(cell_value, 100.0, "Cell 'D4' does not contain the expected value")
        except Exception as e:
            self.fail(f"Accessing cell 'D4' raised an exception: {e}")

        # Test expanding rows beyond existing columns
        self.sheet.expand_sheet_to_include_cell('B6')
        self.assertEqual(len(self.sheet._data.index), 6, "Sheet did not expand to include row 6")

        # Test expanding to a far column and row
        self.sheet.expand_sheet_to_include_cell('Z10')
        # Expect the sheet to have expanded to include column 'Z' and row 10
        self.assertIn('Z', self.sheet._data.columns, "Sheet did not expand to include column 'Z'")
        self.assertEqual(len(self.sheet._data.index), 10, "Sheet did not expand to include row 10")

        # Verify that the cell at 'Z10' can be accessed without error
        try:
            self.sheet.set_cell_value('Z10', '200')
            cell_value = self.sheet.get_cell_value('Z10')
            self.assertEqual(cell_value, 200.0, "Cell 'Z10' does not contain the expected value")
        except Exception as e:
            self.fail(f"Accessing cell 'Z10' raised an exception: {e}")

    def test_expand_sheet_beyond_limits_raises_error(self):
        """Test that an exception is raised when exceeding the maximum"""
        self.sheet.create_sheet(500, 500)  # Initialize the sheet to the maximum size

        # Test expanding beyond the maximum number of columns
        with self.assertRaises(ValueError) as cm:
            self.sheet.expand_sheet_to_include_cell('BAA100')
        self.assertEqual(str(cm.exception), "Exceeding maximum sheet size of 500x500.")

        # Test expanding beyond the maximum number of rows
        with self.assertRaises(ValueError) as cm:
            self.sheet.expand_sheet_to_include_cell('A1001')
        self.assertEqual(str(cm.exception), "Exceeding maximum sheet size of 500x500.")

        # Test expanding beyond both the maximum columns and rows simultaneously
        with self.assertRaises(ValueError) as cm:
            self.sheet.expand_sheet_to_include_cell('ALF1001')
        self.assertEqual(str(cm.exception), "Exceeding maximum sheet size of 500x500.")

    def test_enter_data_in_range(self):
        """Test entering data across a range of cells."""
        self.sheet.create_sheet(5, 5)
        self.sheet.enter_data(("A1", "B2"), "100")

        # Verify that each cell in the range has the entered value
        self.assertEqual(self.sheet.get_cell_value("A1"), 100.0)
        self.assertEqual(self.sheet.get_cell_value("A2"), 100.0)
        self.assertEqual(self.sheet.get_cell_value("B1"), 100.0)
        self.assertEqual(self.sheet.get_cell_value("B2"), 100.0)

    def test_batch_formula_entry(self):
        """Test batch entering formulas into a range of cells."""
        self.sheet.create_sheet(5, 5)
        self.sheet.set_cell_value("A1", "10")
        self.sheet.set_cell_value("B1", "20")
        self.sheet.enter_data(("A2", "B2"), "=A1 + B1")

        self.assertEqual(self.sheet.get_cell_value("A2"), 30.0)
        self.assertEqual(self.sheet.get_cell_value("B2"), 30.0)

    def test_enter_data_with_invalid_cell(self):
        """Test handling of invalid cell ranges in enter_data method."""
        self.sheet.create_sheet(5, 5)
        with self.assertRaises(ValueError):
            # Assuming an invalid range raises a ValueError in the implementation
            self.sheet.enter_data(("A1", "InvalidCell"), "100")

    def test_enter_data_invalid_range_raises_error(self):
        """Test entering data into an invalid cell range raises an error."""
        self.sheet.create_sheet(5, 5)

        # Attempting to enter data into an invalid range where the start cell comes after the end cell
        with self.assertRaises(ValueError) as cm:
            cells_range = ("B2", "A1")
            self.sheet.enter_data(("B2", "A1"), "100")
        self.assertEqual(str(cm.exception), f"Invalid cell range: '{cells_range[0]}' is after '{cells_range[1]}'.")

        # Optionally, test other invalid ranges
        with self.assertRaises(ValueError) as cm:
            cells_range = ("C3", "A2")
            self.sheet.enter_data(("C3", "A2"), "200")
        self.assertEqual(str(cm.exception), f"Invalid cell range: '{cells_range[0]}' is after '{cells_range[1]}'.")

    def test_clear_sheet_resets_values(self):
        """Test that clear_sheet resets all cell values to None."""
        self.sheet.create_sheet(5, 5)
        # Set up some initial values and formulas
        self.sheet.set_cell_value("A1", "100")
        self.sheet.set_cell_value("B2", "=A1*2")

        # Clear the sheet
        self.sheet.clear_sheet()

        # Check that cell values are reset
        self.assertNotEqual(self.sheet.get_cell_value("A1"), 100, "A1 should be reset to None")
        self.assertNotEqual(self.sheet.get_cell_value("B2"), "=A1*2", "B2 should be reset to None")

    def test_clear_sheet_preserves_structure(self):
        """Test that clear_sheet preserves the number of rows and columns."""
        self.sheet.create_sheet(5, 5)
        # Clear the sheet
        self.sheet.clear_sheet()

        # Check that the structure is preserved
        rows, cols = self.sheet._data.shape
        self.assertEqual(rows, 5, "Number of rows should be preserved")
        self.assertEqual(cols, 5, "Number of columns should be preserved")

    def test_clear_sheet_resets_formulas_and_dependencies(self):
        """Test that clear_sheet resets formulas and dependencies."""
        self.sheet.create_sheet(5, 5)
        # Set up some formulas and dependencies
        self.sheet.set_cell_value("A1", "100")
        self.sheet.set_cell_value("B2", "=A1*2")

        # Clear the sheet
        self.sheet.clear_sheet()

        # Check that formulas, functions, and dependencies are cleared
        self.assertEqual(len(self.sheet._formulas), 0, "Formulas should be cleared")
        self.assertEqual(len(self.sheet._functions), 0, "Functions should be cleared")
        self.assertEqual(len(self.sheet._dependencies), 0, "Dependencies should be cleared")
        self.assertEqual(len(self.sheet._reverse_dependencies), 0, "Reverse dependencies should be cleared")
