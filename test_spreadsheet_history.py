import os
from spreadsheet import Spreadsheet, SpreadsheetHistory
import pytest


class TestSpreadsheetHistory:
    """
    Tests for the SpreadsheetHistory class, focusing on the functionality related to state management,
    including undo operations, state saving and recovery, enforcing history limits, and handling cases with
    no available states.
    """
    @pytest.fixture(scope="function")
    def spreadsheet_history(self, tmp_path):  # Use tmp_path for automatic cleanup
        history_dir = tmp_path / "history"
        history_dir.mkdir()
        history = SpreadsheetHistory(str(history_dir))
        yield history

    @pytest.fixture
    def spreadsheet(self):
        return Spreadsheet()

    def test_undo(self, spreadsheet, spreadsheet_history):
        spreadsheet_history.save_current_state(spreadsheet)
        spreadsheet_history.save_current_state(spreadsheet)
        assert len(os.listdir(spreadsheet_history.history_dir)) == 2, "Should have two state files before undo."
        spreadsheet_history.undo()
        assert len(os.listdir(spreadsheet_history.history_dir)) == 1, "Should have one state file after undo."

    def test_save_current_state(self, spreadsheet, spreadsheet_history):
        spreadsheet_history.save_current_state(spreadsheet)
        assert len(os.listdir(spreadsheet_history.history_dir)) == 1, "Should save one state file."

    def test_recover_last_saved_state(self, spreadsheet, spreadsheet_history):
        spreadsheet_history.save_current_state(spreadsheet)
        recovered_spreadsheet = spreadsheet_history.recover_last_saved_state()
        assert recovered_spreadsheet is not None, "Should recover the last saved state."

    def test_enforce_max_history_limit(self, spreadsheet, spreadsheet_history):
        for _ in range(spreadsheet_history.max_history + 5):
            spreadsheet_history.save_current_state(spreadsheet)
        spreadsheet_history.enforce_max_history_limit()
        assert len(os.listdir(spreadsheet_history.history_dir)) == spreadsheet_history.max_history - 1, \
            f"Should enforce the max history limit of {spreadsheet_history.max_history}."

    def test_can_undo_with_no_states(self, spreadsheet_history):
        for file in os.listdir(spreadsheet_history.history_dir):
            os.remove(os.path.join(spreadsheet_history.history_dir, file))
        assert not spreadsheet_history.can_undo(), "Should not be able to undo with no states available."

    def test_undo_with_no_previous_state_raises_error(self, spreadsheet, spreadsheet_history):
        with pytest.raises(Exception, match="No previous states available for undo."):
            spreadsheet_history.undo()

    def test_recover_last_saved_state_with_no_files_returns_none(self, spreadsheet_history):
        for file in os.listdir(spreadsheet_history.history_dir):
            os.remove(os.path.join(spreadsheet_history.history_dir, file))
        result = spreadsheet_history.recover_last_saved_state()
        assert result is None, "Should return None when no history files are found."

