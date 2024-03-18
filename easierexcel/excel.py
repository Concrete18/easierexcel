from logging.handlers import RotatingFileHandler
import logging as lg

from dataclasses import dataclass, fields
from pathlib import Path
import shutil, os, time, openpyxl, zipfile


@dataclass
class Excel:
    """
    Allows retreiving, adding, updating, deleting and formatting cells within Excel.'

    `filename` is the path to the excel file.

    `use_logging` allows disabling all logs when running.

    `log_file` sets the path for logging.

    `log_level` Sets the logging level of this logger (level must be an int or a str).
    """

    filename: str
    use_logging: bool = True
    log_file: str = "logs/excel.log"
    log_level: lg._levelToName = lg.DEBUG

    def __post_init__(self):
        # sets default variables
        self.changes_made = False
        self.backed_up = False

        # creates workbook or raises error
        self.wb = self.workbook_setup(self.filename)

        # logger setup
        self.use_logging = self.use_logging
        datefmt = "%m-%d-%Y %I:%M:%S %p"
        log_formatter = lg.Formatter(
            "%(asctime)s %(levelname)s %(message)s", datefmt=datefmt
        )
        self.logger = lg.getLogger(__name__)
        self.logger.setLevel(self.log_level)  # Log Level
        max_gigs = 2
        # TODO test this
        if self.use_logging:
            if not os.path.exists(self.log_file):
                os.makedirs(os.path.dirname(self.log_file), exist_ok=True)
                with open(self.log_file, "w"):
                    pass
        my_handler = RotatingFileHandler(
            self.log_file,
            maxBytes=max_gigs * 1024 * 1024,
            backupCount=2,
        )
        my_handler.setFormatter(log_formatter)
        self.logger.addHandler(my_handler)

    def __repr__(self):
        string = "Excel("
        for field in fields(self):
            string += f"\n  {field.name}: {getattr(self, field.name)}"
        string += "\n)"
        return string

    def workbook_setup(self, filename):
        """
        ph
        """
        self.file_path = Path(filename)
        try:
            return openpyxl.load_workbook(self.file_path)
        except zipfile.BadZipFile:
            print(f"Error with {self.file_path}")
            response = input("Do you want to restore backup?\n")
            if response in ["yes", "yeah", "y"]:
                # renames current to .old
                os.rename(self.file_path, f"{self.file_path}.old")
                # copies backup and renames to non backup filename
                shutil.copy(f"{self.file_path}.bak", self.file_path)
                # resetup workbook
                self.wb = openpyxl.load_workbook(self.file_path)
            else:  # pragma: no cover
                raise Exception("Excel file is corrupted.")

    def save(
        self,
        use_print: bool = False,
        force_save: bool = False,
        backup: bool = False,
    ):
        """
        Backs up the excel file before saving the changes if `backup` is True.

        It will keep trying to save until it completes in case of permission
        errors caused by the file being open.

        `use_print` determines if info for the saving progress will be printed.

        `force_save` can be used to make sure a save occurs.
        """
        if not self.file_path.exists():
            raise Exception(f"{self.file_path} no longer exists.")
        # only saves if any changes were made or force_save is used
        if self.changes_made or force_save:
            # backups the file before saving.
            if backup:
                if not self.backed_up:
                    backup_path = f"{self.file_path}.bak"
                    shutil.copy(self.file_path, backup_path)
                    self.backed_up = True
                    self.logger.info(f"Excel file backed up")
            # saves the file once it is closed
            if use_print:
                print("\nSaving...")
            try:
                first_run = True
                while self.changes_made:
                    if self.file_path.exists():
                        # tries to save the file
                        try:
                            self.wb.save(self.file_path)
                            self.changes_made = False
                            if use_print:
                                print(f'Save Complete{35*" "}')
                        # catches error caused by excel worksheet being open
                        except PermissionError:  # pragma: no cover
                            if first_run and use_print:
                                msg = "Make sure the excel sheet is closed."
                                print(msg, end="\r")
                            time.sleep(1)
                    else:  # pragma: no cover
                        print("File no longer exists. Save Cancelled")
                        raise Exception(f"{self.file_path} no longer exists.")
                    first_run = False
            except KeyboardInterrupt:  # pragma: no cover
                print(f"Cancelled Save.")
                exit()
        else:
            msg = "Save Skipped due to no changes being made."
            self.logger.info(msg)
            return False

    def open_excel(
        self,
        save: bool = True,
        test: bool = False,
    ):  # pragma: no cover
        """
        Opens the current excel file if it still exists and then exits.

        Saves changes if `save` is True.

        The `test` arg is only used for testing.
        """
        if save:
            self.save(use_print=False)
        if self.file_path.exists():
            if not test:
                os.startfile(self.file_path)
        else:
            raise Exception(f"{self.file_path} no longer exists.")


if __name__ == "__main__":
    excel_file = Excel(filename="test/excel_test.xlsx")
    print(excel_file)
