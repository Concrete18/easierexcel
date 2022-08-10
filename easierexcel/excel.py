from logging.handlers import RotatingFileHandler
import logging as lg
import shutil, os, sys, time, openpyxl, zipfile
from pathlib import Path


class Excel:

    changes_made = False
    backed_up = False
    ext_terminal = sys.stdout.isatty()

    def __init__(
        self,
        filename: str,
        use_logging: bool = True,
        log_file: str = "excel.log",
        log_level=lg.DEBUG,
    ):
        """
        Allows retreiving, adding, updating, deleting and
        formatting cells within Excel.

        `filename` is the path to the excel file.

        `use_logging` allows disabling all logs when running.

        `log_file` sets the path for logging.

        `log_level` Sets the logging level of this logger.
        level must be an int or a str.
        """
        # workbook setup
        self.file_path = Path(filename)
        try:
            self.wb = openpyxl.load_workbook(self.file_path)
        except zipfile.BadZipFile:
            raise Exception("Excel file is currupted.")
            # TODO decide if the below should be used or not
            # testing for this is uncertain
            # print(f"Error with {self.file_path}.")
            # response = input("Do you want to restore backup?")
            # if response in ["yes", "yeah", "y"]:
            #     # renames current to .old
            #     os.rename(self.file_path, f"{self.file_path}.old")
            #     # renames backup to remove .bak
            #     os.rename(f"{self.file_path}.bak", self.file_path)
        # logger setup
        self.use_logging = use_logging
        datefmt = "%m-%d-%Y %I:%M:%S %p"
        log_formatter = lg.Formatter(
            "%(asctime)s %(levelname)s %(message)s", datefmt=datefmt
        )
        self.logger = lg.getLogger(__name__)
        self.logger.setLevel(log_level)  # Log Level
        max_gigs = 2
        my_handler = RotatingFileHandler(
            log_file,
            maxBytes=max_gigs * 1024 * 1024,
            backupCount=2,
        )
        my_handler.setFormatter(log_formatter)
        self.logger.addHandler(my_handler)

    def save(
        self,
        use_print: bool = True,
        force_save: bool = False,
        backup: bool = True,
    ):
        """
        `use_print` determines if info for the saving progress will be printed.

        `force_save` can be used to make sure a save occurs.

        Backs up the excel file before saving the changes if `backup` is True.

        It will keep trying to save until it completes in case of permission
        errors caused by the file being open.
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
                                print(f'Save Complete.{34*" "}')
                        # catches error caused by excel worksheet being open
                        except PermissionError:  # pragma: no cover
                            if first_run and use_print:
                                msg = "Make sure the excel sheet is closed."
                                print(msg, end="\r")
                            time.sleep(1)
                    else:
                        print("File no longer exists. Save Cancelled")
                        raise Exception(f"{self.file_path} no longer exists.")
                    first_run = False
            except KeyboardInterrupt:  # pragma: no cover
                print(f"Cancelled Save.")
                exit()
        else:
            if use_print:
                msg = "Save Skipped due to no changes being made."
                self.logger.info(msg)
                print(msg)
            return False

    def open_excel(
        self,
        save: bool = True,
        exit_after: bool = True,
        test: bool = False,
    ):  # pragma: no cover
        """
        Opens the current excel file if it still exists and then exits.

        Saves changes if `save` is True.
        """
        if save:
            self.save()
        if self.file_path.exists():
            if not test:
                os.startfile(self.file_path)
        else:
            # TODO raise Error
            print("File no longer exists.")
        if exit_after:
            exit()

    def open_file_input(self):  # pragma: no cover
        """
        Opens the excel file if it exists after enter is
        pressed during the input.
        """
        if not self.ext_terminal:
            self.save()
            exit()
        try:
            input("\nPress Enter to open the excel sheet.\n")
        except KeyboardInterrupt:
            print("Closing...")
            self.save()
            exit()
        self.open_excel()
