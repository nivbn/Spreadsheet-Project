import glob
import pickle
import tempfile
import os
from simpleeval import SimpleEval


class SpreadsheetHistory:
    """
    Manages the history of spreadsheet states to enable multiple undo functionality and crash recovery.
    Maintains a limited number of temporary state files.
    """

    def __init__(self, history_dir, max_history=20):
        """
        Initializes the SpreadsheetHistory manager with a directory for storing state files and
        sets a limit on the number of history states to keep.

        Args:
            history_dir: The directory where spreadsheet state files are to be stored.
            max_history (int): The maximum number of history state files to maintain.
        """
        self.history_dir = history_dir
        self.max_history = max_history
        self.states = []  # List to keep track of the state file paths

    @staticmethod
    def save_state(spreadsheet, recovery_file_path: str):
        """
        Saves the current state of the spreadsheet to a temporary file in a specified directory.
        Each saved state allows for recovery and undo functionality.

        Args:
            spreadsheet: The Spreadsheet object to be saved.
            recovery_file_path: The path to the file containing a previously saved state.
        """
        # Temporarily remove the _evaluator before pickling
        evaluator = spreadsheet._evaluator
        spreadsheet._evaluator = None

        # Directly write the state to the specified recovery file
        with open(recovery_file_path, 'wb') as f:
            pickle.dump(spreadsheet, f)

        # Restore the _evaluator after pickling
        spreadsheet._evaluator = evaluator

    @staticmethod
    def load_state(filename: str):
        """
        Loads a spreadsheet state from a file, effectively restoring its saved state.

        Args:
            filename: The path to the file containing a previously saved state.

        Returns:
            A Spreadsheet object restored to its saved state.
        """
        with open(filename, 'rb') as f:
            spreadsheet = pickle.load(f)  # Deserialize the spreadsheet object from the file
        # Reinitialize the SimpleEval instance here
        spreadsheet._evaluator = SimpleEval()
        return spreadsheet  # Return the deserialized Spreadsheet object

    def enforce_max_history_limit(self):
        """
        Ensures that no more than max_history state files exist in the history directory.
        If there are more, the oldest files will be deleted until the limit is met.
        """
        # Get a list of all .pkl files in the history directory
        files = glob.glob(os.path.join(self.history_dir, 'spreadsheet_state_*.pkl'))
        # Sort files by modification time, oldest first
        files.sort(key=os.path.getmtime)

        # If there are more files than max_history, remove the oldest ones
        while len(files) >= self.max_history:
            os.remove(files.pop(0))

    def save_current_state(self, spreadsheet):
        """
        Saves the current state of the spreadsheet to a temporary file, respecting the max history limit.
        Oldest states are deleted as new ones are added beyond the limit.

        Args:
            spreadsheet: The Spreadsheet object whose state is to be saved.
        """
        # Ensure the history directory exists
        if not os.path.exists(self.history_dir):
            os.makedirs(self.history_dir)

        # First, enforce max history limit before saving a new state
        self.enforce_max_history_limit()

        # Create a temporary file name for the new state
        fd, temp_file_path = tempfile.mkstemp(suffix='.pkl', prefix='spreadsheet_state_', dir=self.history_dir)

        # Close the file descriptor immediately to avoid leaks, since we'll open the file by path later
        os.close(fd)

        # Use save_state function to write the state to the new temporary file
        self.save_state(spreadsheet, temp_file_path)

        # Add the new state file path to the history list
        self.states.append(temp_file_path)

    def undo(self):
        """
        Reverts the spreadsheet to its most recent saved state, effectively "undoing" the last operation.
        The reverted state file is deleted to prevent reuse.

        Returns:
            A Spreadsheet object reverted to its previous state, or None if no history is available.
        """
        if not self.states:
            raise Exception("No previous states available for undo.")

        last_state_file = self.states.pop()
        try:
            spreadsheet = self.load_state(last_state_file)
        except Exception as e:
            raise Exception(f"Failed to load the previous state: {e}")

        if spreadsheet is None:
            raise Exception("Failed to undo to the previous state. The state file could not be loaded correctly.")

        # Clean up by removing the state file that was just undone
        os.remove(last_state_file)

        return spreadsheet

    def can_undo(self) -> bool:
        """
        Checks if there are any actions available to undo.

        Returns:
            bool: True if there is at least one state to undo, False otherwise.
        """
        # Count the number of pickle files in the history directory
        pickle_files = glob.glob(os.path.join(self.history_dir, '*.pkl'))

        # Ensure there are actions to undo and more than two state file available
        return len(self.states) > 0 and len(pickle_files) > 1

    def recover_last_saved_state(self):
        """
        Recovers the spreadsheet from the last saved pickle file in the history_files directory.

        Returns:
            A Spreadsheet object recovered from the most recent saved state, or None if no history files are found.
        """
        # Get a list of all .pkl files in the history directory
        files = glob.glob(os.path.join(self.history_dir, '*.pkl'))
        if not files:
            return None

        # Find the most recent file based on the modification time
        latest_file = max(files, key=os.path.getmtime)

        # Load the state from the most recent file
        recovered_spreadsheet = self.load_state(latest_file)

        return recovered_spreadsheet
