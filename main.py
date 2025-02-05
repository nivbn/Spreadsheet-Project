from spreadsheet_GUI import *
import sys
import os


def print_help_file(help_file_path):
    """Prints the contents of the help text file to the console."""
    try:
        with open(help_file_path, 'r') as file:
            print(file.read())
    except FileNotFoundError:
        print("Help file not found. Try to find manually in the project's directory")
    except Exception as e:
        print(f"Error reading help file: {e}. Try to find manually in the project's directory")


# Main application setup
if __name__ == "__main__":
    # Construct an absolute path to the help file
    dir_path = os.path.dirname(os.path.realpath(__file__))
    help_file_path = os.path.join(dir_path, "help.txt")

    # Check if --help argument was provided in the command line arguments
    if "--help" in sys.argv:
        # Print the help text file
        print_help_file(help_file_path)
    else:
        # Proceed with the rest of the script to initialize and display the program
        root = tk.Tk()
        root.deiconify()
        root.geometry("1000x600")
        app = SpreadsheetGUI(root)
        root.mainloop()
