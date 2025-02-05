import tkinter as tk
from tkinter import filedialog, ttk, Menu, messagebox
from typing import Optional
import tkinter.font as tkFont
from tkinter.simpledialog import Dialog
from tkintertable import TableCanvas, TableModel, ColumnHeader
from spreadsheet import Spreadsheet


class SpreadsheetGUI:
    """
        A GUI representation of a spreadsheet program.

        Attributes:
            master (tk.Tk): The Tkinter root window.
            _spreadsheet (Spreadsheet): The Spreadsheet object for data management.
            fill_start_cell (bool): Toggle determining which cell to fill on click.
            menu_bar (Menu): The menu bar for the GUI.
            cell_entry (tk.Entry): Entry widget for start cell address.
            end_cell_entry (tk.Entry): Entry widget for end cell address.
            value_entry (tk.Entry): Entry widget for cell value or formula.
            operation_var (tk.StringVar): Variable storing selected operation.
            operation_combobox (ttk.Combobox): Combobox for selecting operations.
        """

    def __init__(self, master: tk.Tk, rows: int = 25, columns: int = 10) -> None:
        """
        Initializes the Spreadsheet GUI with a given number of rows and columns.
        """
        self.master = master
        self._spreadsheet = Spreadsheet()
        self._spreadsheet.create_sheet(rows, columns)

        # Initialize the GUI components
        master.title("Spreadsheet Program")
        self.setup_menu()
        self.setup_operation_frame()
        self.setup_table()
        self.update_table()

        # Set the start cell to be filled on the first click by default
        self.fill_start_cell = True

        # Initialize the last clicked entry to the start cell entry
        self.last_clicked_entry = self.cell_entry

        # Bind click events for entries to update last_clicked_entry
        self.cell_entry.bind('<FocusIn>', lambda e: self.set_last_clicked_entry(self.cell_entry))
        self.end_cell_entry.bind('<FocusIn>', lambda e: self.set_last_clicked_entry(self.end_cell_entry))
        self.value_entry.bind('<FocusIn>', lambda e: self.set_last_clicked_entry(self.value_entry))

    def setup_menu(self) -> None:
        """
        Sets up the menu bar with file operations.
        """
        self.menu_bar = Menu(self.master)
        self.master.config(menu=self.menu_bar)

        # File menu
        self.file_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.file_menu.add_command(label="Recover", command=self.recover_sheet)
        self.file_menu.add_command(label="Load", command=self.load_sheet)
        self.file_menu.add_command(label="Save", command=self.save_sheet)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Clear Sheet", command=self.clear_sheet)

        # Edit menu
        self.edit_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Edit", menu=self.edit_menu)
        self.edit_menu.add_command(label="Column Width", command=self.prompt_column_width)

        # Help menu
        self.help_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
        self.help_menu.add_command(label="About", command=self.show_help_dialog)

    def setup_operation_frame(self) -> None:
        """Sets up the operation frame with cell address entries and operation selection."""
        self.create_operation_frame()
        self.setup_left_side_container()
        self.setup_operation_combobox()
        self.setup_cell_entries()
        self.setup_value_entry()
        self.setup_execute_button()
        self.setup_right_side_container()
        self.setup_undo_button()
        self.setup_cell_value_display()
        self.bind_return_key_to_execute_command()

    def create_operation_frame(self):
        """Creates the main operation frame."""
        self.operation_frame = tk.Frame(self.master)
        self.operation_frame.pack(fill=tk.X, side=tk.TOP)

    def setup_left_side_container(self):
        """Creates the left side container for entries and buttons."""
        self.left_side_container = tk.Frame(self.operation_frame)
        self.left_side_container.pack(side=tk.LEFT, expand=False)

    def setup_operation_combobox(self):
        """Sets up the operation selection combobox."""
        operations = ["Set Value", "Sum", "Average", "Max", "Min", "Count", "Median", "Product", "Delete Value"]
        self.operation_var = tk.StringVar()
        self.operation_combobox = ttk.Combobox(self.left_side_container, textvariable=self.operation_var,
                                               values=operations, width=10)
        self.operation_combobox.pack(side=tk.LEFT, padx=(10, 0))
        self.operation_combobox.set("Set Value")
        self.operation_combobox.bind("<<ComboboxSelected>>", self.update_value_entry_title)

    def setup_cell_entries(self):
        """Sets up the start and end cell entries."""
        start_cell_title = tk.Label(self.left_side_container, text="Start Cell:")
        start_cell_title.pack(side=tk.LEFT, padx=(10, 0))
        self.cell_entry = tk.Entry(self.left_side_container, width=10)
        self.cell_entry.pack(side=tk.LEFT, padx=(10, 0))

        end_cell_title = tk.Label(self.left_side_container, text="End Cell:")
        end_cell_title.pack(side=tk.LEFT, padx=(10, 0))
        self.end_cell_entry = tk.Entry(self.left_side_container, width=10)
        self.end_cell_entry.pack(side=tk.LEFT, padx=(10, 0))

    def setup_value_entry(self):
        """Sets up the value/formula entry."""
        self.value_entry_title_var = tk.StringVar(value="Value:")
        value_entry_title = tk.Label(self.left_side_container, textvariable=self.value_entry_title_var)
        value_entry_title.pack(side=tk.LEFT, padx=(10, 0))
        self.value_entry = tk.Entry(self.left_side_container, width=15)
        self.value_entry.pack(side=tk.LEFT, padx=(10, 0))

    def setup_execute_button(self):
        """Sets up the execute command button."""
        self.execute_button = tk.Button(self.left_side_container, text="Enter", command=self.execute_command)
        self.execute_button.pack(side=tk.LEFT, padx=(10, 0))

    def setup_right_side_container(self):
        """Creates the right side container for the undo button and cell data display."""
        self.right_side_container = tk.Frame(self.operation_frame)
        self.right_side_container.pack(side=tk.RIGHT, fill=tk.X)

    def setup_undo_button(self):
        """Sets up the undo button."""
        self.undo_button = tk.Button(self.right_side_container, text="Undo", command=self.undo)
        self.undo_button.pack(side=tk.RIGHT, padx=(10, 0))
        self.update_undo_button_state()

    def setup_cell_value_display(self):
        """Sets up the display for cell values."""
        display_font = tkFont.Font(size=10, weight='bold')
        start_cell_title = tk.Label(self.right_side_container, text="Cell Data:", font=display_font)
        start_cell_title.pack(side=tk.LEFT, padx=(10, 0))
        self.cell_value_display = tk.Label(self.right_side_container, text="", relief="sunken", anchor="e",
                                           font=display_font)
        self.cell_value_display.pack(fill=tk.BOTH, expand=True, padx=(10, 0))

    def bind_return_key_to_execute_command(self):
        """Binds the Return key to the execute command action for cell and value entries."""
        self.cell_entry.bind("<Return>", self.execute_command)
        self.end_cell_entry.bind("<Return>", self.execute_command)
        self.value_entry.bind("<Return>", self.execute_command)

    def setup_table(self) -> None:
        """
        Configures and displays the table component for the spreadsheet data.
        """
        self.table_frame = tk.Frame(self.master)
        self.table_frame.pack(fill=tk.BOTH, expand=True, side=tk.BOTTOM)
        self.table_model = TableModel()
        self.table = TableCanvas(self.table_frame, model=self.table_model, read_only=True,
                                 selectedcolor='Yellow', rowselectedcolor='white')
        self.table_column_header = ColumnHeader(table=self.table)
        self.table.show()
        self.table.autoResizeColumns()

        # Bind table click to custom handler
        self.table.bind('<Button-1>', self.on_table_click)

        self.disable_bindings()

    def disable_bindings(self) -> None:
        """
        Disables specific event bindings for the spreadsheet GUI.
        This method unbinds mouse and keyboard events that are not required or could interfere with the intended
        operation of the spreadsheet GUI.
        """
        if hasattr(self, 'main_widget'):  # Check if the main_widget attribute exists
            # Unbind mouse drag behavior on the main widget
            self.main_widget.unbind('<B1-Motion>')

            # Unbind mouse movement behavior on the main widget
            self.main_widget.unbind('<Motion>')

        # Disable keys
        self.table.unbind('<B1-Motion>')
        self.table.unbind_all('<Delete>')
        self.table.unbind_all("<Control-x>")
        self.table.unbind_all("<Control-n>")
        self.table.unbind_all("<Control-v>")
        self.table.unbind_all("<Shift>")
        self.table.unbind_all("<Control>")
        self.table.unbind_all("<Left>")
        self.table.unbind_all("<Right>")
        self.table.unbind_all("<Up>")
        self.table.unbind_all("<Down>")
        self.table.unbind("<Control-Button-1>")
        self.table.unbind("<Shift-Button-1>")

        # Prevent column dragging
        self.disable_column_dragging()

    def disable_column_dragging(self) -> None:
        """
        Disables column dragging in the table view.
        """
        column_header = self.table.tablecolheader
        if column_header:
            column_header.unbind('<Button-1>')
            column_header.unbind('<B1-Motion>')
            column_header.unbind('<ButtonRelease-1>')
            column_header.unbind('<Motion>')

    def update_value_entry_title(self, event: Optional[tk.Event] = None) -> None:
        """
        Updates the title of the value entry based on the selected operation.

        Args:
            event (Optional[tk.Event]): The event that triggered the update. Default is None.
        """
        operation = self.operation_var.get()
        if operation == "Set Value" or operation == "Delete Value":
            self.value_entry_title_var.set("Value:")
        else:
            self.value_entry_title_var.set("Target Cell:")

        # Reset the contents of the entry widgets
        self.cell_entry.delete(0, tk.END)
        self.end_cell_entry.delete(0, tk.END)
        self.value_entry.delete(0, tk.END)

        # Set focus to the cell entry
        self.cell_entry.focus_set()

    def on_table_click(self, event: tk.Event) -> None:
        """
        Handles table click events, updating UI components with cell information.

        Args:
            event (tk.Event): The mouse click event.
        """
        # Determine clicked row and column
        row, col = self.table.get_row_clicked(event), self.table.get_col_clicked(event)

        if row is not None and col is not None:
            # Convert row and column to a cell address (e.g., A1)
            cell_address = self._spreadsheet.convert_indices_to_cell_address(row, col)

            # Highlight the selected cell
            self.table.setSelectedRow(row)
            self.table.setSelectedCol(col)
            self.table.redrawTable()

            # Initialize display_value as empty
            display_value = ""

            # Check if the cell has an associated function
            if cell_address in self._spreadsheet._functions:
                func_name, cell_range = self._spreadsheet._functions[cell_address]
                # Format the display value as a function (e.g., =SUM(A1:A4))
                display_value = f"={func_name}({cell_range[0]}:{cell_range[1]})"

            # If no function, check for formula or direct value
            if not display_value:
                cell_formula = self._spreadsheet.get_cell_formula(cell_address)
                cell_value = self._spreadsheet.get_cell_value(cell_address)

                if cell_formula is not None:
                    display_value = "=" + cell_formula
                elif cell_value:
                    display_value = str(cell_value)  # Ensure it's a string
                    if display_value == "nan":
                        display_value = ""

            # Update the GUI to show the function/formula/value
            self.cell_value_display.config(text=display_value)

            # Update cell address entries
            self.update_cell_entries(cell_address)

    def set_last_clicked_entry(self, entry_widget: tk.Entry) -> None:
        """
        Sets the last clicked entry widget for cell address update.

        Args:
            entry_widget (tk.Entry): The entry widget that was clicked.
        """
        self.last_clicked_entry = entry_widget

    def update_cell_entries(self, cell_address: str) -> None:
        """
        Updates cell address entries after a cell is clicked or selected.

        Specifically, if the value_entry is the last clicked entry, each additional click on a cell appends the cell
        address to the existing content in the value_entry, ensuring it starts with "=".

        Args:
            cell_address (str): The address of the clicked or selected cell.
        """
        if self.last_clicked_entry:
            # Check if the last clicked entry is the value_entry
            if self.last_clicked_entry == self.value_entry and self.value_entry_title_var.get() == "Value:":
                # Retrieve the current content from value_entry
                current_content = self.value_entry.get()

                # If there's already content, and it starts with "=", append the new cell address
                # Otherwise, start a new content with "="
                if current_content and current_content.startswith("="):
                    new_content = current_content + cell_address
                else:
                    new_content = "=" + cell_address

                # Update value_entry with the new content
                self.value_entry.delete(0, tk.END)
                self.value_entry.insert(0, new_content)
            else:
                # For other entries, replace the content with the new cell address
                if self.last_clicked_entry is self.cell_entry:
                    self.last_clicked_entry.delete(0, tk.END)
                    self.last_clicked_entry.insert(0, cell_address)
                    self.end_cell_entry.delete(0, tk.END)
                    self.end_cell_entry.insert(0, cell_address)
                else:
                    self.last_clicked_entry.delete(0, tk.END)
                    self.last_clicked_entry.insert(0, cell_address)

            # Set focus back to the last clicked entry for immediate input
            self.last_clicked_entry.focus_set()

    def handle_operation(self, operation: str, start_cell: str, end_cell: str, target_or_value: str) -> None:
        """
        Handles _spreadsheet operations based on user input.

        Args:
            operation (str): The operation to perform.
            start_cell (str): The starting cell address.
            end_cell (str): The ending cell address.
            target_or_value (str): The value or target cell for the operation.
        """
        # Handling the "Set value" and "Delete value" operation differently
        if operation == "Set Value":
            if start_cell and not end_cell:  # If only a start cell is provided
                self._spreadsheet.set_cell_value(start_cell, target_or_value)
            elif end_cell and not start_cell:  # If only end cell is provided
                self._spreadsheet.set_cell_value(end_cell, target_or_value)
            elif start_cell and end_cell:  # If both start and end cells are provided
                self._spreadsheet.enter_data((start_cell, end_cell), target_or_value)

        elif operation == "Delete Value":
            if start_cell and not end_cell:  # If only a start cell is provided
                self._spreadsheet.set_cell_value(start_cell, "")
            elif end_cell and not start_cell:  # If only end cell is provided
                self._spreadsheet.set_cell_value(end_cell, "")
            elif start_cell and end_cell:  # If both start and end cells are provided
                self._spreadsheet.enter_data((start_cell, end_cell), "")

        else:
            # Use target_or_value as the target cell for the operation result
            cell_range = (start_cell, end_cell) if end_cell else start_cell
            self._spreadsheet.execute_function(target_or_value, operation, cell_range)

    def execute_command(self, event: Optional[tk.Event] = None) -> None:
        """
        Executes the selected operation on the spreadsheet based on the user's input.
        """
        operation = self.operation_var.get()
        start_cell = self.cell_entry.get().strip()
        end_cell = self.end_cell_entry.get().strip()
        target_or_value = self.value_entry.get().strip()

        try:
            # Save recovery file before execution
            self._spreadsheet.history_manager.save_current_state(self._spreadsheet)

            # Handle operation logic
            self.handle_operation(operation, start_cell, end_cell, target_or_value)

            self.undo_button.config(state=tk.NORMAL)
            self.update_table()  # Refresh the table display
            self.clear_input_fields()  # Clear input fields after operation execution
            self.cell_entry.focus_set()

        except Exception as e:
            # Display the error message in a messagebox
            self.update_table()  # Refresh the table display
            messagebox.showerror("Error", str(e), parent=self.master)

    def update_table(self) -> None:
        """
        Updates the table display with the current data from the _spreadsheet model.
        """
        # Create a temporary DataFrame for display purposes
        temp_df = self._spreadsheet._data.copy()

        # Replace 'nan' values with an empty string '' in the temporary DataFrame
        temp_df = temp_df.fillna('')

        # Ensure all values are strings, including booleans and numerics
        temp_df = temp_df.map(str)

        # Convert the temporary DataFrame to a format that can be used by TableModel
        data = {'rec' + str(i): row.to_dict() for i, row in temp_df.iterrows()}

        self.table_model.deleteRows()  # Clear existing data
        self.table_model.importDict(data)  # Load new data
        self.table.redraw()  # Redraw table
        self.table.autoResizeColumns()  # Auto-resize the columns to fit new content
        self.clear_input_fields()

    def clear_input_fields(self) -> None:
        """
        Clears the input fields after an operation is executed.
        """
        self.cell_entry.delete(0, tk.END)
        self.end_cell_entry.delete(0, tk.END)
        self.value_entry.delete(0, tk.END)

    def prompt_column_width(self) -> None:
        """
        Prompts the user to set the width for a column in the spreadsheet. The prompt defaults
        to the last clicked column or to column 'A' if no column was clicked before invoking this method.
        """
        # Determine the last clicked column. If no column was clicked, default to 'A'
        last_clicked_col_index = self.table.getSelectedColumn()
        last_clicked_col = self.table.model.getColumnName(
            last_clicked_col_index) if last_clicked_col_index >= 0 else 'A'

        # Default width is the width of the last clicked column or column A
        default_width = self.table.model.columnwidths.get(last_clicked_col, 100)

        # Pass the last clicked column as an argument to SizeDialog
        dialog = SizeDialog(self.master, "Column Width", default_width, last_clicked_col)
        if dialog.result:
            column, size = dialog.result
            self.set_column_width(column, size)

    def set_column_width(self, column: str, size: str) -> None:
        """
        Sets the width of a specified column.

        Args:
            column (str): The column identifier, which can be a letter or name.
            size (str): The new width for the column, as a string to be converted to an integer.
        """
        # Convert column letter to uppercase to ensure consistency
        column = column.upper()

        # Convert the size to an integer and ensure it's a positive value up to 500
        try:
            size = int(size)
            if size <= 0 or size > 500:
                raise ValueError("Size must be a positive integer up to 500.")
        except ValueError:
            messagebox.showerror("Error", "Invalid size. Size must be a positive integer up to 500.")
            return

        # Attempt to update the column width
        try:
            if column in self.table.model.columnNames:
                self.table.model.columnwidths[column] = size
            else:
                raise ValueError("Invalid column identifier.")
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return

        # Redraw the table to reflect the new column width
        self.table.redrawTable()

    def undo(self) -> None:
        """
         Reverts the last change made to the spreadsheet, if possible.
         """
        if self._spreadsheet.history_manager.can_undo():
            self._spreadsheet = self._spreadsheet.history_manager.undo()
            self.update_table()
        # Update the state of the undo button after undoing
        self.update_undo_button_state()

    def update_undo_button_state(self) -> None:
        """
        Updates the "Undo" button's state based on the availability of undo actions.
        """
        if self._spreadsheet.history_manager.can_undo():
            self.undo_button.config(state=tk.NORMAL)
        else:
            self.undo_button.config(state=tk.DISABLED)

    def clear_sheet(self) -> None:
        """
        Clears all data from the _spreadsheet after confirmation from the user.
        """
        if messagebox.askyesno("Confirm", "Are you sure you want to delete the sheet data?"):
            self._spreadsheet.clear_sheet()
            self.update_table()
            messagebox.showinfo("Info", "Sheet cleared")

    def load_sheet(self) -> None:
        """
        Opens a file dialog for the user to select a _spreadsheet file to load.
        """
        filename = filedialog.askopenfilename(title="Open Spreadsheet", filetypes=[("YAML files", "*.yaml")])
        if filename:
            try:
                self._spreadsheet.load_sheet(filename)
                self.update_table()
            except Exception as e:
                messagebox.showerror("Loading Error", f"Failed to load the spreadsheet: {e}")

    def save_sheet(self) -> None:
        """
        Opens a save file dialog for the user to save the current _spreadsheet.
        """
        filename = filedialog.asksaveasfilename(title="Save Spreadsheet", filetypes=[("YAML files", "*.yaml")],
                                                defaultextension=".yaml")
        if filename:
            self._spreadsheet.save_sheet(filename)

    def recover_sheet(self) -> None:
        """
        Recovers the _spreadsheet from the last pickle file
        """
        recovered_spreadsheet = self._spreadsheet.history_manager.recover_last_saved_state()
        if recovered_spreadsheet:
            self._spreadsheet = recovered_spreadsheet
            self.update_table()
        else:
            messagebox.showerror("Recovery Error", "No recovery files found in the directory")

    def show_help_dialog(self) -> None:
        """
        Displays a help dialog with information about how to use the program.
        """
        help_window = tk.Toplevel(self.master)
        help_window.title("Help and Information")

        # Adjust the size and position of the help window as needed
        help_window.geometry("800x400")

        # Create a scrolling text area for help text
        text_scroll = tk.Scrollbar(help_window)
        text_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        help_text = tk.Text(help_window, yscrollcommand=text_scroll.set, wrap=tk.WORD)
        help_text.pack(expand=True, fill=tk.BOTH)

        # Load help content from a text file
        with open("help.txt", "r") as file:
            help_content = file.read()

        # Insert help text into the text widget
        help_text.insert(tk.END, help_content)

        # Make the help text read-only
        help_text.config(state=tk.DISABLED)

        # Configure the scrollbar
        text_scroll.config(command=help_text.yview)

        # Add a Close button to the help window
        close_button = tk.Button(help_window, text="Close", command=help_window.destroy)
        close_button.pack(pady=10)

        # Focus on the help window
        help_window.focus_set()


class SizeDialog(Dialog):
    def __init__(self, parent: tk.Tk, title: str, default_value: int, default_column: str) -> None:
        """
        Initializes a dialog for setting the size of a column.

        Args:
            parent: The parent window for the dialog.
            title (str): The title of the dialog window.
            default_value (int): The default size value to display in the dialog.
            default_column (str): The default column identifier to display in the dialog.
        """
        self.default_value = default_value
        self.default_column = default_column  # Default column identifier
        super().__init__(parent, title)

    def body(self, master) -> tk.Entry:
        """
        Creates the body of the dialog.
        """
        # Label and entry for specifying the column (e.g., "Column: A")
        tk.Label(master, text="Column:").grid(row=0)
        self.column_entry = tk.Entry(master)
        self.column_entry.insert(0, self.default_column)  # Pre-fill with the default column identifier
        self.column_entry.grid(row=0, column=1)

        # Label and entry for specifying the size (e.g., "Size: 100")
        tk.Label(master, text="Size:").grid(row=1)
        self.size_entry = tk.Entry(master)
        self.size_entry.insert(0, str(self.default_value))  # Pre-fill with the default size value
        self.size_entry.grid(row=1, column=1)

        return self.column_entry  # Set initial focus on the column entry

    def apply(self) -> None:
        """
        Processes the input when the user submits the dialog, setting the result attribute.
        """
        # The result attribute is set to a tuple containing the entered column identifier and size value.
        self.result = (self.column_entry.get(), self.size_entry.get())
