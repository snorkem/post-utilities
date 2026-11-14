from __future__ import annotations

import os
import pandas as pd
from pathlib import Path
import sys
import re
import logging
import argparse
from datetime import datetime
from dataclasses import dataclass
from typing import Iterable, List, Tuple, Optional


@dataclass
class RenameConfig:
    """Configuration for file renaming operations."""
    suffix: str = "_M"
    force_overwrite: bool = False
    dry_run: bool = False


@dataclass
class RenameResult:
    """Results of a batch rename operation."""
    renamed_count: int = 0
    skipped_count: int = 0
    errors_count: int = 0

    def get_summary(self, dry_run: bool = False) -> str:
        """Generate a human-readable summary."""
        prefix = "Dry run - " if dry_run else ""
        verb = "would be " if dry_run else ""
        return (f"{prefix}Summary: {self.renamed_count} files {verb}renamed, "
                f"{self.skipped_count} skipped (destinations exist), {self.errors_count} errors")


class InputError(Exception):
    """A problem with the inputs that the user has to fix; the message is shown to them."""

    def __init__(self, message: str, title: str = "Input Error"):
        super().__init__(message)
        self.title = title


class FileValidator:
    """Validates filenames and detects issues."""

    # Path separators, characters Windows forbids in names, and control characters
    INVALID_CHARS_PATTERN = re.compile(r'[<>:"/\\|?*{}\[\]~`\x00-\x1f]')

    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or logging.getLogger(__name__)

    def validate_filename(self, filename, index: int, column_name: str) -> bool:
        """Check if filename is valid ascii characters and has no illegal path characters."""
        try:
            if pd.isna(filename) or not str(filename).strip():
                error_msg = f'Empty cell in {column_name} on row {index+2}'
                print(error_msg)
                self.logger.error(error_msg)
                return False

            if self.INVALID_CHARS_PATTERN.search(str(filename)):
                error_msg = f'Invalid path char detected in {column_name} on row {index+2}'
                print(error_msg)
                self.logger.error(error_msg)
                return False

            # Try to encode to ascii to catch non-ascii characters
            str(filename).encode('ascii')
            return True
        except UnicodeEncodeError:
            error_msg = f'Non-ascii name in {column_name} "{filename}" on row {index+2}'
            print(error_msg)
            self.logger.error(error_msg)
            return False
        except AttributeError:
            error_msg = f'Empty cell, perhaps, on line {index+2}?'
            print(error_msg)
            self.logger.error(error_msg)
            return False

    def check_duplicates(self, series: pd.Series, column_name: str) -> bool:
        """Check if a pandas Series contains duplicate values."""
        if not series.is_unique:
            error_msg = f"Duplicate values found in '{column_name}' column"
            print(error_msg)
            self.logger.error(error_msg)
            return True
        return False


class FileCollector:
    """Collects files from directories based on criteria."""

    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or logging.getLogger(__name__)
        self.unreadable_dirs: List[Path] = []

    def collect_files(self, target_dir: Path, recursive: bool = True) -> List[Path]:
        """Collect non-hidden files, skipping hidden folders (e.g. .Trashes, .git).

        Folders that can't be read are skipped and recorded in unreadable_dirs,
        so one locked folder doesn't stop the rest from being collected.
        """
        files_to_rename = []

        if not target_dir.is_dir():
            print('Target is not a directory')
            self.logger.error('Target is not a directory: %s', target_dir)
            return files_to_rename

        print(f'Target is directory: {target_dir}. Collecting files...')
        self.logger.info('Target is directory: %s. Collecting files...', target_dir)

        def skip_unreadable(error: OSError) -> None:
            unreadable = Path(error.filename)
            self.unreadable_dirs.append(unreadable)
            error_msg = f'Skipping folder that could not be read: {unreadable} ({error.strerror})'
            print(error_msg)
            self.logger.warning(error_msg)

        # os.walk rather than Path.walk (Python 3.12+) to keep older Pythons working
        for dir_path, dir_names, file_names in os.walk(target_dir, onerror=skip_unreadable):
            dir_names[:] = sorted(d for d in dir_names if not d.startswith('.'))
            files_to_rename.extend(Path(dir_path) / name
                                   for name in sorted(file_names) if not name.startswith('.'))
            if not recursive:
                break

        return files_to_rename

    def collect_from_dirs(self, target_dirs: Iterable[Path], recursive: bool = True) -> List[Path]:
        """Collect files from several directories, listing each file only once
        even when the directories overlap (e.g. a folder and its subfolder)."""
        seen = set()
        collected = []
        for target_dir in target_dirs:
            dir_files = self.collect_files(target_dir, recursive)
            new_files = [f for f in dir_files if f.resolve() not in seen]
            seen.update(f.resolve() for f in new_files)
            collected.extend(new_files)
            print(f"Found {len(dir_files)} files in {target_dir}"
                  + (f" ({len(dir_files) - len(new_files)} already collected)" if len(new_files) < len(dir_files) else ""))
        return collected


class FileRenamer:
    """Handles file renaming operations with safety features."""

    def __init__(self, config: RenameConfig, logger: Optional[logging.Logger] = None):
        self.config = config
        self.logger = logger or logging.getLogger(__name__)

    def rename_files(self, target_files: List[Path], master_file_names: pd.Series,
                     proxy_names: pd.Series) -> RenameResult:
        """Rename files based on mapping in Excel file."""
        result = RenameResult()

        # Create a lookup dictionary for faster access
        rename_map = {}
        for index, master_name in master_file_names.items():
            if pd.notna(master_name) and pd.notna(proxy_names[index]):
                rename_map[str(master_name)] = str(proxy_names[index])

        print(f'Examining {len(target_files)} files...')
        self.logger.info(f'Examining {len(target_files)} files...')

        # Process each file
        for file_path in target_files:
            target_name = rename_map.get(file_path.stem)
            if target_name is None:
                continue

            new_name = target_name + self.config.suffix + file_path.suffix
            new_path = file_path.parent / new_name

            try:
                if new_path == file_path:
                    print(f"Skipping {file_path} - already has the new name")
                    self.logger.info(f"Skipping rename - already named: {file_path}")
                    result.skipped_count += 1
                    continue

                destination_exists = new_path.exists()
                if destination_exists and not self.config.force_overwrite:
                    print(f"Skipping {file_path} - destination already exists: {new_path}")
                    self.logger.warning(f"Skipping rename - destination exists: {new_path}")
                    result.skipped_count += 1
                    continue

                # Simulate, reporting the same actions a real run would take
                if self.config.dry_run:
                    if destination_exists:
                        print(f"Would back up existing {new_path.name} and overwrite it")
                        self.logger.info(f"Dry run: would back up and overwrite {new_path}")
                    print(f"Would rename: {file_path.name} -> {new_name}")
                    self.logger.info(f"Dry run: would rename {file_path} -> {new_path}")
                    result.renamed_count += 1
                    continue

                backup_path = None
                if destination_exists:
                    # Move the existing file aside rather than copying it: instant even for
                    # huge media files, and it frees the name (Windows won't rename onto
                    # an existing file)
                    backup_path = self._backup_path(new_path)
                    new_path.rename(backup_path)
                    print(f"Backed up existing file to: {backup_path}")
                    self.logger.info(f"Backed up existing file: {new_path} -> {backup_path}")

                try:
                    file_path.rename(new_path)
                except OSError:
                    if backup_path is not None:
                        # Put the original destination file back before reporting the error
                        backup_path.rename(new_path)
                        self.logger.info(f"Restored {new_path} from backup after failed rename")
                    raise

                print(f"Renamed: {file_path.name} -> {new_path.name}")
                self.logger.info(f"Renamed: {file_path} -> {new_path}")
                result.renamed_count += 1

            except Exception as e:
                error_msg = f"Error processing {file_path}: {str(e)}"
                print(error_msg)
                self.logger.error(error_msg)
                result.errors_count += 1

        summary = result.get_summary(self.config.dry_run)
        print(summary)
        self.logger.info(summary)

        return result

    @staticmethod
    def _backup_path(path: Path) -> Path:
        """Unused name for a backup of path, e.g. 'B001_M_backup_20251027140501.mov'."""
        stamp = datetime.now().strftime('%Y%m%d%H%M%S')
        candidate = path.with_name(f"{path.stem}_backup_{stamp}{path.suffix}")
        counter = 1
        while candidate.exists():
            candidate = path.with_name(f"{path.stem}_backup_{stamp}_{counter}{path.suffix}")
            counter += 1
        return candidate


class ExcelReader:
    """Reads and validates Excel files for file renaming operations."""

    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or logging.getLogger(__name__)

    def check_single_sheet(self, excel_file: Path) -> Tuple[bool, Optional[str]]:
        """Check if the Excel file contains only a single sheet."""
        try:
            # Use pandas ExcelFile to get sheet names (closed afterwards so the file isn't left locked)
            with pd.ExcelFile(excel_file) as xls:
                sheet_names = xls.sheet_names

            if len(sheet_names) > 1:
                error_msg = f"Excel file contains multiple sheets: {', '.join(sheet_names)}. Only single-sheet files are supported."
                print(error_msg)
                self.logger.error(error_msg)
                return False, sheet_names[0]

            return True, sheet_names[0]

        except Exception as e:
            error_msg = f"Error checking Excel sheets: {str(e)}"
            print(error_msg)
            self.logger.error(error_msg)
            return False, None


def _strip_cells(series: pd.Series) -> pd.Series:
    """Strip stray whitespace from text cells; empty and numeric cells are left as they are."""
    return series.map(lambda value: value.strip() if isinstance(value, str) else value)


def load_name_mapping(excel_file: Path, files_col: str, rename_col: str,
                      logger: logging.Logger) -> Tuple[pd.Series, pd.Series]:
    """Read and validate the original and new filename columns from the spreadsheet.

    Returns (original names, new names). Raises InputError describing the first
    problem found, so the CLI and GUI report the same checks the same way.
    """
    excel_reader = ExcelReader(logger)
    validator = FileValidator(logger)

    # Check if Excel file has only one sheet
    single_sheet, sheet_name = excel_reader.check_single_sheet(excel_file)
    if sheet_name is None:
        raise InputError("The Excel file could not be read.", "Excel Error")
    if not single_sheet:
        raise InputError("This tool only supports Excel files with a single sheet.", "Excel Error")

    # Read the Excel file into dataframe
    try:
        df = pd.read_excel(excel_file, sheet_name=0)  # Always use first sheet
    except Exception as e:
        raise InputError(f"Error reading Excel file: {str(e)}", "Excel Error") from e
    print(f"Successfully read Excel file with {len(df)} rows from sheet '{sheet_name}'")
    logger.info(f"Successfully read Excel file with {len(df)} rows from sheet '{sheet_name}'")

    # Verify required columns exist
    for column in (files_col, rename_col):
        if column not in df.columns:
            raise InputError(f"Required column not found: '{column}'", "Column Error")

    master_file_names = _strip_cells(df[files_col])
    proxy_names = _strip_cells(df[rename_col])

    # Validate filenames
    invalid_rows = []
    for index, (master_name, proxy_name) in enumerate(zip(master_file_names, proxy_names)):
        if not validator.validate_filename(master_name, index, files_col) or \
           not validator.validate_filename(proxy_name, index, rename_col):
            invalid_rows.append(index + 2)  # +2 for Excel row number

    if invalid_rows:
        raise InputError(f"Invalid filenames found in rows: {', '.join(map(str, invalid_rows))}",
                         "Validation Error")

    # Check for duplicates
    for series, column in ((master_file_names, files_col), (proxy_names, rename_col)):
        if validator.check_duplicates(series, column):
            raise InputError(f"Duplicate values found in '{column}' column", "Duplicate Error")

    return master_file_names, proxy_names


def setup_logger(logname='renames.log'):
    """Set up and return a logger with appropriate configuration."""
    log_file_path = str(Path(logname).resolve())

    logging.basicConfig(
        filename=log_file_path,
        filemode='a',
        format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
        datefmt="%Y-%m-%d %H:%M:%S",
        level=logging.DEBUG
    )
    return logging.getLogger('Rename Files From Excel'), log_file_path


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Rename files based on Excel spreadsheet mapping.',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument('excel_file',
                        help='Path to Excel file containing filename mappings')

    parser.add_argument('target_dir',
                        help='Directory containing files to rename')

    parser.add_argument('--files-col',
                        default='Master Original Name',
                        help='Column name in Excel containing original filenames')

    parser.add_argument('--rename-col',
                        default='Avid Proxy Name',
                        help='Column name in Excel containing new filenames')

    parser.add_argument('--suffix',
                        default='_M',
                        help='Suffix to add to renamed files')

    parser.add_argument('--log',
                        default='renames.log',
                        help='Path to log file')

    parser.add_argument('--no-recursive',
                        action='store_true',
                        dest='no_recursive',
                        default=False,
                        help='Disable recursive search in subdirectories (default: recursive search enabled)')

    parser.add_argument('--force',
                        action='store_true',
                        help='Overwrite existing files during rename (existing files are kept as backups)')

    parser.add_argument('--dry-run',
                        action='store_true',
                        help='Show what would be renamed without making changes')

    return parser.parse_args()


def main():
    # Parse command line arguments
    args = parse_arguments()

    # Set up logger
    logger, log_path = setup_logger(args.log)
    print(f"Logging to: {log_path}")

    # Convert string paths to absolute Path objects
    excel_file = Path(args.excel_file).resolve()
    target_dir = Path(args.target_dir).resolve()

    # Column names from arguments
    files_to_rename_col = args.files_col
    rename_to_col = args.rename_col
    master_file_suffix = args.suffix

    # Determine if recursive search is enabled (default is True)
    recursive = not args.no_recursive

    print(f"Starting file renaming process:")
    print(f"Excel file: {excel_file}")
    print(f"Target directory: {target_dir}")
    print(f"Original filenames column: {files_to_rename_col}")
    print(f"New filenames column: {rename_to_col}")
    print(f"Suffix: {master_file_suffix}")
    print(f"Recursive search: {recursive}")
    print(f"Force overwrite: {args.force}")
    print(f"Dry run: {args.dry_run}")
    print("-" * 50)

    # Log the configuration
    logger.info("Configuration: excel=%s, target_dir=%s, files_col=%s, rename_col=%s, suffix=%s, recursive=%s, force=%s, dry_run=%s",
               excel_file, target_dir, files_to_rename_col, rename_to_col, master_file_suffix,
               recursive, args.force, args.dry_run)

    try:
        # Create component instances
        collector = FileCollector(logger)
        rename_config = RenameConfig(
            suffix=master_file_suffix,
            force_overwrite=args.force,
            dry_run=args.dry_run
        )
        renamer = FileRenamer(rename_config, logger)

        # Verify the Excel file has the correct extension
        if not excel_file.name.lower().endswith(('.xlsx', '.xls')):
            raise InputError(f"The specified file is not an Excel file (.xlsx or .xls): {excel_file}")

        # Verify the Excel file exists
        if not excel_file.exists():
            raise InputError(f"The specified Excel file does not exist: {excel_file}")

        master_file_names, proxy_names = load_name_mapping(
            excel_file, files_to_rename_col, rename_to_col, logger)

        # Verify target directory exists
        if not target_dir.exists():
            raise InputError(f"Target directory does not exist: {target_dir}")

        if not target_dir.is_dir():
            raise InputError(f"Target is not a directory: {target_dir}")

        # Collect files to process
        files_to_rename = collector.collect_files(target_dir, recursive)

        if not files_to_rename:
            print("No files found to rename.")
            logger.info("No files found to rename.")
            sys.exit(0)

        # Perform renaming
        renamer.rename_files(files_to_rename, master_file_names, proxy_names)

    except InputError as e:
        print(str(e))
        logger.error(str(e))
        sys.exit(1)

    except Exception as e:
        error_msg = f"Error: {str(e)}"
        print(error_msg)
        logger.exception(error_msg)
        sys.exit(1)


if __name__ == '__main__':
    main()
