from __future__ import annotations

import os
import sys
from pathlib import Path
import threading
import traceback
import subprocess
import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                            QLabel, QLineEdit, QPushButton, QListWidget, QTextEdit,
                            QGroupBox, QGridLayout, QFileDialog, QMessageBox, QStatusBar,
                            QScrollArea, QCheckBox)
from PyQt5.QtCore import Qt, QObject, pyqtSignal, pyqtSlot

# Import shared classes from CLI module
from rename_files_from_excel import (
    FileCollector, FileRenamer, InputError, RenameConfig, load_name_mapping, setup_logger
)

# Keep the log next to this script, not wherever the app happened to be launched from
LOG_PATH = Path(__file__).resolve().parent / 'renames.log'


class TextRedirector(QObject):
    """Class to redirect stdout to a QTextEdit widget."""
    text_written = pyqtSignal(str)

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.text_written.connect(self.update_text)
        self.buffer = ""

    def write(self, string):
        self.buffer += string
        self.text_written.emit(string)

    def flush(self):
        pass

    @pyqtSlot(str)
    def update_text(self, text):
        self.text_widget.moveCursor(self.text_widget.textCursor().End)
        self.text_widget.insertPlainText(text)
        self.text_widget.ensureCursorVisible()


class FileRenamerApp(QMainWindow):
    """GUI application for renaming files based on Excel data."""

    # The renaming runs on a worker thread, which must never touch widgets.
    # It emits these instead; Qt delivers them to the slots on the main thread.
    message_requested = pyqtSignal(str, str, str)  # level, title, text
    processing_finished = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Renamer Tool")
        self.setMinimumSize(800, 800)
        self.resize(800, 1200)

        # Column names in Excel
        self.master_name_col = "Master Original Name"
        self.proxy_name_col = "Avid Proxy Name"
        self.master_suffix = "_M"

        # File and directory paths
        self.excel_path = ""
        self.target_dirs = []

        # Processing options
        self.recursive_search = True
        self.force_overwrite = False
        self.dry_run = False

        # Setup logger using shared function
        self.logger, self.log_file_path = setup_logger(str(LOG_PATH))

        # Create UI components
        self.init_ui()

        # Store backup of stdout for redirection
        self.stdout_backup = sys.stdout

        self.message_requested.connect(self.show_message)
        self.processing_finished.connect(self.on_processing_finished)

    def init_ui(self):
        """Initialize the user interface."""
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Brief explanation at the top
        explanation_frame = QGroupBox()
        explanation_layout = QVBoxLayout(explanation_frame)
        explanation_text = "This tool renames files based on an Excel spreadsheet. It looks for files with names matching the 'Master Original Name' column\nand renames them to the corresponding name in the 'Avid Proxy Name' column, adding the suffix at the end."
        explanation_label = QLabel(explanation_text)
        explanation_label.setAlignment(Qt.AlignLeft)
        explanation_layout.addWidget(explanation_label)
        main_layout.addWidget(explanation_frame)

        # Excel file selection section
        excel_frame = QGroupBox("Excel File Selection")
        excel_layout = QGridLayout(excel_frame)

        excel_label = QLabel("Excel File:")
        excel_layout.addWidget(excel_label, 0, 0)

        self.excel_path_edit = QLineEdit()
        excel_layout.addWidget(self.excel_path_edit, 0, 1)

        excel_browse_btn = QPushButton("Browse...")
        excel_browse_btn.clicked.connect(self.browse_excel_file)
        excel_layout.addWidget(excel_browse_btn, 0, 2)

        main_layout.addWidget(excel_frame)

        # Directory selection section
        dir_frame = QGroupBox("Target Directories")
        dir_layout = QVBoxLayout(dir_frame)

        dir_button_frame = QWidget()
        dir_button_layout = QHBoxLayout(dir_button_frame)

        dir_label = QLabel("Directory:")
        dir_button_layout.addWidget(dir_label)

        self.dir_entry = QLineEdit()
        dir_button_layout.addWidget(self.dir_entry)

        dir_browse_btn = QPushButton("Browse...")
        dir_browse_btn.clicked.connect(self.browse_directory)
        dir_button_layout.addWidget(dir_browse_btn)

        add_dir_btn = QPushButton("Add Directory")
        add_dir_btn.clicked.connect(self.add_directory)
        dir_button_layout.addWidget(add_dir_btn)

        remove_dir_btn = QPushButton("Remove Selected")
        remove_dir_btn.clicked.connect(self.remove_directory)
        dir_button_layout.addWidget(remove_dir_btn)

        dir_layout.addWidget(dir_button_frame)

        # Directory listbox
        self.dir_listbox = QListWidget()
        dir_layout.addWidget(self.dir_listbox)

        main_layout.addWidget(dir_frame)

        # Excel column configuration section
        config_frame = QGroupBox("Excel Configuration")
        config_layout = QGridLayout(config_frame)

        master_label = QLabel("Master Original Name Column:")
        config_layout.addWidget(master_label, 0, 0)

        self.master_col_edit = QLineEdit(self.master_name_col)
        config_layout.addWidget(self.master_col_edit, 0, 1)

        master_desc = QLabel("(Column containing original filenames to search for)")
        config_layout.addWidget(master_desc, 0, 2)

        proxy_label = QLabel("Avid Proxy Name Column:")
        config_layout.addWidget(proxy_label, 1, 0)

        self.proxy_col_edit = QLineEdit(self.proxy_name_col)
        config_layout.addWidget(self.proxy_col_edit, 1, 1)

        proxy_desc = QLabel("(Column containing new filenames to rename to)")
        config_layout.addWidget(proxy_desc, 1, 2)

        suffix_label = QLabel("Master File Suffix:")
        config_layout.addWidget(suffix_label, 2, 0)

        self.suffix_edit = QLineEdit(self.master_suffix)
        config_layout.addWidget(self.suffix_edit, 2, 1)

        suffix_desc = QLabel("(Added to the end of each renamed file)")
        config_layout.addWidget(suffix_desc, 2, 2)

        main_layout.addWidget(config_frame)

        # Options section
        options_frame = QGroupBox("Options")
        options_layout = QVBoxLayout(options_frame)

        self.recursive_checkbox = QCheckBox("Search files recursively in subdirectories")
        self.recursive_checkbox.setChecked(self.recursive_search)
        options_layout.addWidget(self.recursive_checkbox)

        self.force_checkbox = QCheckBox("Force overwrite (existing destination files are kept as backups)")
        self.force_checkbox.setChecked(self.force_overwrite)
        options_layout.addWidget(self.force_checkbox)

        self.dry_run_checkbox = QCheckBox("Dry run (show what would be renamed without making changes)")
        self.dry_run_checkbox.setChecked(self.dry_run)
        options_layout.addWidget(self.dry_run_checkbox)

        main_layout.addWidget(options_frame)

        # Action buttons
        button_frame = QWidget()
        button_layout = QHBoxLayout(button_frame)

        log_button = QPushButton("View Log File")
        log_button.clicked.connect(self.open_log_file)
        button_layout.addWidget(log_button)

        button_layout.addStretch()

        self.rename_button = QPushButton("RENAME FILES")
        self.rename_button.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.rename_button.clicked.connect(self.start_renaming)
        button_layout.addWidget(self.rename_button)

        main_layout.addWidget(button_frame)

        # Output console
        console_frame = QGroupBox("Console Output")
        console_layout = QVBoxLayout(console_frame)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        console_layout.addWidget(self.console)

        main_layout.addWidget(console_frame)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def browse_excel_file(self):
        """Browse for Excel file and update the path field."""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")

            if file_path:
                self.excel_path_edit.setText(file_path)
                self.status_bar.showMessage(f"Excel file selected: {Path(file_path).name}")

        except Exception as e:
            error_msg = f"Error browsing for file: {str(e)}"
            self.logger.error(error_msg)
            QMessageBox.critical(self, "Error", f"Could not browse for file: {str(e)}")

    def browse_directory(self):
        """Browse for directory and update the path field."""
        try:
            dir_path = QFileDialog.getExistingDirectory(
                self, "Select Target Directory", "")

            if dir_path:
                self.dir_entry.setText(dir_path)
                self.status_bar.showMessage(f"Directory selected: {dir_path}")

        except Exception as e:
            error_msg = f"Error browsing for directory: {str(e)}"
            self.logger.error(error_msg)
            QMessageBox.critical(self, "Error", f"Could not browse for directory: {str(e)}")

    def open_log_file(self):
        """Open the log file in the default system text editor."""
        try:
            log_path = Path(self.log_file_path)
            if not log_path.exists():
                # Create an empty log file if it doesn't exist yet
                log_path.write_text("Log file created on {}\n".format(
                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

            # Open the log file with the system's default text editor
            if sys.platform == "win32":
                # Windows
                os.startfile(self.log_file_path)
            elif sys.platform == "darwin":
                # macOS
                subprocess.Popen(["open", self.log_file_path])
            else:
                # Linux and other Unix-like
                subprocess.Popen(["xdg-open", self.log_file_path])

            self.status_bar.showMessage(f"Opened log file: {self.log_file_path}")
        except Exception as e:
            error_msg = f"Error opening log file: {str(e)}"
            self.logger.error(error_msg)
            QMessageBox.critical(self, "Error", error_msg)

            # Fallback: Show the log file path so the user can find it manually
            QMessageBox.information(self, "Log File Location",
                               f"The log file is located at:\n{self.log_file_path}\n\n"
                               f"Please open it manually with a text editor.")

    def add_directory(self):
        """Add a directory to the list."""
        dir_path = self.dir_entry.text().strip()
        if not dir_path:
            QMessageBox.information(self, "Input Required", "Please enter a directory path.")
            return

        if dir_path not in self.target_dirs:
            self.target_dirs.append(dir_path)
            self.dir_listbox.addItem(dir_path)
            self.dir_entry.clear()  # Clear the entry
            self.status_bar.showMessage(f"Added directory: {dir_path}")
        else:
            QMessageBox.warning(self, "Duplicate Directory", "This directory is already in the list.")

    def remove_directory(self):
        """Remove selected directory from the list."""
        selected_items = self.dir_listbox.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Selection Required", "Please select a directory to remove.")
            return

        for item in selected_items:
            row = self.dir_listbox.row(item)
            self.dir_listbox.takeItem(row)
            self.target_dirs.remove(item.text())

        self.status_bar.showMessage("Directory removed")

    def validate_inputs(self):
        """Quick checks on the inputs before starting; the spreadsheet itself is
        read and validated on the worker thread."""
        excel_path = self.excel_path_edit.text()
        if not excel_path:
            QMessageBox.critical(self, "Input Error", "Please enter an Excel file path.")
            return False

        # Verify the Excel file has the correct extension
        if not excel_path.lower().endswith(('.xlsx', '.xls')):
            QMessageBox.critical(self, "Input Error", "The specified file is not an Excel file (.xlsx or .xls).")
            return False

        if not Path(excel_path).exists():
            QMessageBox.critical(self, "Input Error", "The specified Excel file does not exist.")
            return False

        if not self.target_dirs:
            QMessageBox.critical(self, "Input Error", "Please add at least one target directory.")
            return False

        for dir_path in self.target_dirs:
            if not Path(dir_path).exists():
                QMessageBox.critical(self, "Input Error", f"Directory does not exist: {dir_path}")
                return False

        if not self.master_col_edit.text() or not self.proxy_col_edit.text():
            QMessageBox.critical(self, "Input Error", "Please provide column names for Master and Proxy names.")
            return False

        return True

    def start_renaming(self):
        """Start the file renaming process in a separate thread."""
        if not self.validate_inputs():
            return

        # Update configuration values from UI
        self.master_name_col = self.master_col_edit.text()
        self.proxy_name_col = self.proxy_col_edit.text()
        self.master_suffix = self.suffix_edit.text()
        self.excel_path = self.excel_path_edit.text()
        self.recursive_search = self.recursive_checkbox.isChecked()
        self.force_overwrite = self.force_checkbox.isChecked()
        self.dry_run = self.dry_run_checkbox.isChecked()

        # Disable UI during processing
        self.disable_ui()
        self.console.clear()

        # Redirect stdout to console
        sys.stdout = TextRedirector(self.console)

        # Start processing in a separate thread
        thread = threading.Thread(target=self.rename_files_process)
        thread.daemon = True
        thread.start()

    def disable_ui(self):
        """Disable UI controls during processing."""
        central_widget = self.centralWidget()
        for child in central_widget.findChildren(QPushButton):
            child.setEnabled(False)

        for child in central_widget.findChildren(QLineEdit):
            child.setEnabled(False)

        for child in central_widget.findChildren(QCheckBox):
            child.setEnabled(False)

        self.dir_listbox.setEnabled(False)
        self.status_bar.showMessage("Processing... Please wait.")

    def enable_ui(self):
        """Re-enable UI controls after processing."""
        central_widget = self.centralWidget()
        for child in central_widget.findChildren(QPushButton):
            child.setEnabled(True)

        for child in central_widget.findChildren(QLineEdit):
            child.setEnabled(True)

        for child in central_widget.findChildren(QCheckBox):
            child.setEnabled(True)

        self.dir_listbox.setEnabled(True)
        self.status_bar.showMessage("Ready")

    @pyqtSlot(str, str, str)
    def show_message(self, level, title, text):
        """Show a message box (main thread only; the worker requests these via message_requested)."""
        {'critical': QMessageBox.critical,
         'warning': QMessageBox.warning}.get(level, QMessageBox.information)(self, title, text)

    @pyqtSlot()
    def on_processing_finished(self):
        """Restore stdout and the controls once the worker thread is done."""
        sys.stdout = self.stdout_backup
        self.enable_ui()

    def rename_files_process(self):
        """Main file renaming process (runs on a worker thread).

        Never touch widgets here: report through message_requested and
        processing_finished, which run on the main thread.
        """
        try:
            # Get configuration values
            excel_file = Path(self.excel_path).resolve()
            target_dirs = [Path(dir_path).resolve() for dir_path in self.target_dirs]
            master_col = self.master_name_col
            proxy_col = self.proxy_name_col
            suffix = self.master_suffix
            recursive = self.recursive_search
            force = self.force_overwrite
            dry_run = self.dry_run

            print(f"Starting file renaming process:")
            print(f"Excel file: {excel_file}")
            print(f"Target directories: {', '.join(str(d) for d in target_dirs)}")
            print(f"Master column: {master_col}")
            print(f"Proxy column: {proxy_col}")
            print(f"Suffix: {suffix}")
            print(f"Recursive search: {recursive}")
            print(f"Force overwrite: {force}")
            print(f"Dry run: {dry_run}")
            print("-" * 50)

            # Log the configuration
            self.logger.info("Configuration: excel=%s, dirs=%s, master_col=%s, proxy_col=%s, "
                           "suffix=%s, recursive=%s, force=%s, dry_run=%s",
                           excel_file, target_dirs, master_col, proxy_col,
                           suffix, recursive, force, dry_run)

            # Create shared component instances
            collector = FileCollector(self.logger)
            rename_config = RenameConfig(suffix=suffix, force_overwrite=force, dry_run=dry_run)
            renamer = FileRenamer(rename_config, self.logger)

            # Read and validate the spreadsheet (same checks as the CLI)
            try:
                master_file_names, proxy_names = load_name_mapping(
                    excel_file, master_col, proxy_col, self.logger)
            except InputError as e:
                print(str(e))
                self.logger.error(str(e))
                self.message_requested.emit('critical', e.title, str(e))
                return

            # Collect files once each, even if the chosen folders overlap
            files_to_rename = collector.collect_from_dirs(target_dirs, recursive)
            print(f"Total files collected from all directories: {len(files_to_rename)}")

            if not files_to_rename:
                print("No files found to rename.")
                self.logger.info("No files found to rename.")
                self.message_requested.emit('information', "No Files",
                                            "No files found to rename in the specified directories.")
                return

            # Perform renaming using shared renamer
            result = renamer.rename_files(files_to_rename, master_file_names, proxy_names)

            # Show completion message with every outcome, not just the successes
            completion_msg = result.get_summary(dry_run)
            if collector.unreadable_dirs:
                completion_msg += f"\n{len(collector.unreadable_dirs)} folders could not be read and were skipped."
            completion_msg += f"\n\nSee the log for details:\n{self.log_file_path}"
            print(f"See '{self.log_file_path}' for complete details.")
            self.logger.info(completion_msg)
            problems = result.errors_count or collector.unreadable_dirs
            self.message_requested.emit('warning' if problems else 'information',
                                        "Process Complete", completion_msg)

        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}\n{traceback.format_exc()}"
            print(error_msg)
            self.logger.exception(error_msg)
            self.message_requested.emit('critical', "Error", error_msg)

        finally:
            # Restore stdout and re-enable the UI on the main thread
            self.processing_finished.emit()


def main():
    # Set up error handler for uncaught exceptions
    def show_error(exc_type, exc_value, exc_traceback):
        error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        print(error_msg)  # Print to console/terminal

        # Try to show error dialog if QApplication is still working
        try:
            QMessageBox.critical(None, 'Unhandled Exception',
                               f"An error occurred: {str(exc_value)}\n\nSee log for details.")
        except:
            pass  # If messagebox fails, at least we printed to console

    # Set the exception hook
    sys.excepthook = show_error

    # Create and run the application
    app = QApplication(sys.argv)
    window = FileRenamerApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
