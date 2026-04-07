"""Worker threads for background operations."""
import re
import time
import sqlite3
import logging
import numpy as np
from typing import Dict
import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal
from xlsxwriter.workbook import Workbook

from utils.file_handler import FileHandler
from utils.data_processor import DataProcessor
from config.constants import DATABASE_PATH, PREVALIDATION_SQL_FILE, REQUIRED_DIRS

logger = logging.getLogger(__name__)


class FileUploadWorker(QThread):
    """Worker thread for handling Excel file uploads without freezing the UI."""

    file_loaded = pyqtSignal(dict)
    progress_update = pyqtSignal(int)
    error_occurred = pyqtSignal(str)

    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            if not self.file_path:
                self.error_occurred.emit("No file selected")
                return

            self.progress_update.emit(25)
            logger.info(f"Starting file upload: {self.file_path}")

            file_name = FileHandler.get_file_name(self.file_path)
            self.progress_update.emit(50)

            df, elapsed = DataProcessor.read_excel_file(self.file_path)
            if df is None:
                self.error_occurred.emit(f"Failed to read file: {file_name}")
                return

            self.progress_update.emit(100)
            logger.info(f"File loaded successfully: {file_name} ({len(df)} rows)")
            self.file_loaded.emit({
                "dataframe": df,
                "filename": file_name,
                "read_time": elapsed,
            })

        except Exception as e:
            logger.exception("Error in file upload worker")
            self.error_occurred.emit(str(e))


class PrevalidationWorker(QThread):
    """Worker thread for data prevalidation, SQL prep, and warning generation."""

    worker_complete = pyqtSignal(dict)
    worker_loading = pyqtSignal(int)

    def __init__(self, source_df: pd.DataFrame, target_df: pd.DataFrame):
        super().__init__()
        self.df = source_df.copy()
        self.df1 = target_df.copy()

    def run(self):
        try:
            self.worker_loading.emit(1)
            error_message = "--ERROR--\n"
            t1 = time.time()

            # Always change first column to string
            self.df.iloc[:, 0] = self.df.iloc[:, 0].astype(str)
            self.df1.iloc[:, 0] = self.df1.iloc[:, 0].astype(str)

            # Store original first column names
            src_changecolA = self.df.columns[0]
            tgt_changecolA = self.df1.columns[0]

            # Identify date columns
            dateCol = [col for col in self.df.select_dtypes(include=[np.datetime64]).columns]
            dateCol1 = [col for col in self.df1.select_dtypes(include=[np.datetime64]).columns]

            # Clean dataframes — convert nullable ints, strip whitespace
            self.df = self._clean_dataframe(self.df)
            self.df1 = self._clean_dataframe(self.df1)

            # Check for special characters in column names
            error_message = self._check_special_chars(error_message)

            # Check column name lengths
            error_message = self._check_column_lengths(error_message)

            # Convert columns with None & Int back to Int64
            self._convert_int_columns(self.df, dateCol)
            self._convert_int_columns(self.df1, dateCol1)

            # Revert pandas ".1" suffix on duplicate column names
            self.df.columns = self.df.columns.str.split('.').str[0]
            self.df1.columns = self.df1.columns.str.split('.').str[0]

            # Check for duplicated column names
            error_message = self._check_duplicate_columns(error_message)

            t2 = time.time()
            if len(error_message) != 10:  # More than just "--ERROR--\n"
                error_message += "Please fix the error(s) and reupload the files."
                self.worker_loading.emit(2)
                self.worker_complete.emit({
                    "df": self.df, "df1": self.df1,
                    "errorMessage": error_message, "time": t2 - t1,
                    "src_changecolA": src_changecolA, "tgt_changecolA": tgt_changecolA,
                })
                return

            # --- Warning phase ---
            warning_count = 0
            inconsistent_dict = self._detect_mixed_datatypes()
            if inconsistent_dict['source']:
                warning_count += 1
            if inconsistent_dict['target']:
                warning_count += 1

            # Prepare copies for SQL (rename first col to ColumnA)
            df_changed = self.df.copy()
            df1_changed = self.df1.copy()

            sqlFormat = self._build_sql_format(df_changed, dateCol, rename_first=True)
            sqlFormat1 = self._build_sql_format(df1_changed, dateCol1, rename_first=True)

            # Update the copies with renamed first column
            df_changed.rename(columns={df_changed.columns[0]: "ColumnA"}, inplace=True)
            df1_changed.rename(columns={df1_changed.columns[0]: "ColumnA"}, inplace=True)

            # Database operations
            warning_count += self._run_database_validation(
                df_changed, df1_changed, sqlFormat, sqlFormat1,
                dateCol, dateCol1, inconsistent_dict,
            )

            t2 = time.time()
            self.worker_loading.emit(2)
            self.worker_complete.emit({
                "df": self.df, "df1": self.df1,
                "warningCount": warning_count, "time": t2 - t1,
                "src_changecolA": src_changecolA, "tgt_changecolA": tgt_changecolA,
            })

        except Exception as e:
            logger.exception("Error in prevalidation worker")
            self.worker_loading.emit(2)
            self.worker_complete.emit({
                "df": self.df, "df1": self.df1,
                "errorMessage": f"--ERROR--\nUnexpected error: {e}\nPlease fix the error(s) and reupload the files.",
                "time": time.time(),
                "src_changecolA": self.df.columns[0],
                "tgt_changecolA": self.df1.columns[0],
            })

    # ---- Helper methods ----

    @staticmethod
    def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
        for col in df.columns:
            if pd.api.types.is_float_dtype(df[col]):
                if df[col].fillna(0).apply(float.is_integer).all():
                    df[col] = df[col].astype('Int64')
            elif pd.api.types.is_object_dtype(df[col]):
                df[col] = df[col].str.strip()
        return df

    def _check_special_chars(self, error_message: str) -> str:
        for label, frame in [("Source", self.df), ("Target", self.df1)]:
            header = f"{label} Column Name(s) with special characters:\n"
            found = [x for x in frame.columns if re.search(r"\W", x)]
            if found:
                error_message += header + "\n".join(found) + "\n\n"
        return error_message

    def _check_column_lengths(self, error_message: str, max_len: int = 100) -> str:
        for label, frame in [("Source", self.df), ("Target", self.df1)]:
            long_cols = [x for x in frame.columns if len(x) > max_len]
            if long_cols:
                error_message += f"{label} Column Name(s) with more than {max_len} characters:\n"
                error_message += "\n".join(long_cols) + "\n\n"
        return error_message

    @staticmethod
    def _convert_int_columns(df: pd.DataFrame, date_cols: list):
        for col in df.columns:
            if pd.api.types.is_integer_dtype(df[col]) or pd.api.types.is_float_dtype(df[col]):
                if col not in date_cols:
                    df[col] = df[col].astype('Int64', errors='ignore')

    def _check_duplicate_columns(self, error_message: str) -> str:
        for label, frame in [("Source", self.df), ("Target", self.df1)]:
            dupes = frame.columns[frame.columns.duplicated()].tolist()
            if dupes:
                error_message += f"Duplicated Column Name(s) in {label}:\n" + ", ".join(dupes) + "\n\n"
        return error_message

    def _detect_mixed_datatypes(self) -> Dict[str, list]:
        result = {'source': [], 'target': []}
        for key, frame in [('source', self.df), ('target', self.df1)]:
            for col in frame.columns:
                weird = (frame[[col]].map(type) != frame[[col]].iloc[0].apply(type)).any(axis=1)
                if len(frame[weird]) > 0:
                    result[key].append(col)
        return result

    @staticmethod
    def _build_sql_format(df: pd.DataFrame, date_cols: list, rename_first: bool = False) -> str:
        sql_format = ""
        for i, col in enumerate(df.columns):
            if pd.api.types.is_object_dtype(df[col]):
                sql_format += col + " text, "
            elif col in date_cols:
                sql_format += col + " date, "
            else:
                sql_format += col + " real, "
        return sql_format[:-2] if sql_format else ""

    @staticmethod
    def _run_database_validation(df_changed, df1_changed, sql_format, sql_format1,
                                  date_cols, date_cols1, inconsistent_dict) -> int:
        import os
        warning_count = 0
        conn = sqlite3.connect(DATABASE_PATH)
        c = conn.cursor()

        try:
            c.execute("DROP TABLE IF EXISTS SOURCE")
            c.execute("DROP TABLE IF EXISTS TARGET")
            c.execute(f"CREATE TABLE SOURCE ({sql_format})")
            c.execute(f"CREATE TABLE TARGET ({sql_format1})")
            conn.commit()

            df_changed.to_sql('SOURCE', conn, if_exists='replace', index=False)
            df1_changed.to_sql('TARGET', conn, if_exists='replace', index=False)
            logger.info("SOURCE & TARGET DATAFRAME CONVERTED TO DATABASE TABLE")

            # Convert TIMESTAMP columns to DATE only
            c.execute("PRAGMA table_info('SOURCE')")
            for row in c.fetchall():
                if row[2] == "TIMESTAMP":
                    c.execute(f"UPDATE SOURCE SET {row[1]} = DATE({row[1]})")
            c.execute("PRAGMA table_info('TARGET')")
            for row in c.fetchall():
                if row[2] == "TIMESTAMP":
                    c.execute(f"UPDATE TARGET SET {row[1]} = DATE({row[1]})")

            # Run prevalidation SQL
            sql_content = FileHandler.read_sql_file(PREVALIDATION_SQL_FILE)
            if sql_content:
                c.execute("CREATE INDEX IF NOT EXISTS IDX_COLUMNA_SOURCE ON SOURCE(COLUMNA)")
                c.execute("CREATE INDEX IF NOT EXISTS IDX_COLUMNA_TARGET ON TARGET(COLUMNA)")
                c.execute("DROP TABLE IF EXISTS R_PREVALIDATION_OUTPUT_TBL")
                c.execute("CREATE TABLE R_PREVALIDATION_OUTPUT_TBL (UNIQUE_ID INT, DESC TEXT, VAL TEXT)")

                for i, cmd in enumerate(sql_content.split(";")):
                    cmd = cmd.strip()
                    if cmd:
                        try:
                            c.executescript(cmd)
                            logger.info(f"PREVALIDATION CHECK #{i + 1}")
                        except Exception as e:
                            logger.warning(f"Prevalidation check #{i + 1} failed: {e}")
                conn.commit()

            # Write warning report
            c.execute("SELECT DESC, VAL FROM R_PREVALIDATION_OUTPUT_TBL")
            rows = c.fetchall()

            warning_dir = REQUIRED_DIRS.get("warnings", "output/warnings")
            os.makedirs(warning_dir, exist_ok=True)
            warning_path = os.path.join(warning_dir, "warningMessage.xlsx")

            workbook = Workbook(warning_path)
            worksheet = workbook.add_worksheet(name="descriptions")
            bold = workbook.add_format({'bold': True})
            worksheet.write(0, 0, "Description", bold)
            worksheet.write(0, 1, "Value", bold)

            counter = 1
            for row in rows:
                worksheet.write_row(counter, 0, row)
                warning_count += 1
                counter += 1

            for label_key in ['source', 'target']:
                desc = f"Mixed Column(s) with words & numbers in {label_key.title()}:"
                for col_name in inconsistent_dict.get(label_key, []):
                    worksheet.write(counter, 0, desc)
                    worksheet.write(counter, 1, col_name)
                    counter += 1

            workbook.close()

        except Exception as e:
            logger.exception("Error during database validation")
            raise
        finally:
            conn.close()

        return warning_count
