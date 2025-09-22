# standard library
import shutil, os, time, zipfile
from pathlib import Path

# third-party imports
from openpyxl import load_workbook


class Excel:
    """
    Allows retreiving, adding, updating, deleting and formatting cells within Excel.
    """

    def __init__(self, filename: str) -> None:
        """
        `filename` is the path to the excel file.
        """
        self.changes_made = False
        self.backed_up = False
        # workbook setup
        self.file_path = Path(filename)
        try:
            # TODO test using blank filter and see if data_only fixes it
            self.wb = load_workbook(self.file_path, data_only=True)
        except zipfile.BadZipFile:
            print(f"Error with {self.file_path}")
            response = input("Do you want to restore backup?\n")
            if response in ["yes", "yeah", "y"]:
                # renames current to .old
                os.rename(self.file_path, f"{self.file_path}.old")
                # copies backup and renames to non backup filename
                shutil.copy(f"{self.file_path}.bak", self.file_path)
                # sets up workbook with restored backup
                self.wb = load_workbook(self.file_path, data_only=True)
            else:  # pragma: no cover
                raise Exception("Excel file is corrupted.")

    def __repr__(self) -> str:
        return f'Excel(filename="{self.file_path}", Sheets="{self.wb.sheetnames}")'

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
            # saves the file once it is closed
            if use_print:
                print("\nSaving...")
            try:
                first_run = True
                while self.changes_made:
                    if self.file_path.exists():
                        # tries to save the file
                        try:
                            if self.wb:
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
            return False

    def open_excel(
        self,
        save: bool = True,
        test: bool = False,
    ) -> None:  # pragma: no cover
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
    excel_file = Excel(filename="tests/excel_test.xlsx")
    print(excel_file)
