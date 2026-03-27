"""Data processing utilities for Excel file handling."""
import re
import time
import logging
from typing import Tuple, Dict, Optional
import pandas as pd
import numpy as np

logger = logging.getLogger(__name__)


class DataProcessor:
    """Handles data processing operations for Excel files."""
    
    @staticmethod
    def read_excel_file(file_path: str, header: int = 0, sheet_name: int = 0) -> Tuple[Optional[pd.DataFrame], float]:
        """
        Read Excel file and return DataFrame.
        
        Args:
            file_path (str): Path to Excel file
            header (int): Row number to use as column names
            sheet_name (int): Sheet number to read
        
        Returns:
            Tuple[Optional[pd.DataFrame], float]: DataFrame and read time in seconds
        """
        start_time = time.time()
        try:
            df = pd.read_excel(file_path, header=header, sheet_name=sheet_name, engine='openpyxl')
            elapsed_time = time.time() - start_time
            logger.info(f"Excel file read successfully in {elapsed_time:.2f}s: {file_path}")
            return df, elapsed_time
        except Exception as e:
            logger.error(f"Error reading Excel file {file_path}: {str(e)}")
            return None, time.time() - start_time
    
    @staticmethod
    def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean and prepare DataFrame for processing.
        
        Args:
            df (pd.DataFrame): Input DataFrame
        
        Returns:
            pd.DataFrame: Cleaned DataFrame
        """
        df = df.copy()
        
        # Convert first column to string
        df.iloc[:, 0] = df.iloc[:, 0].astype(str)
        
        # Convert int columns that have NaN back to Int64
        for col in df.columns:
            dt = df[col].dtype
            if dt == float:
                if df[col].fillna(0).apply(float.is_integer).all():
                    df[col] = df[col].astype('Int64')
            elif dt == object:
                # Strip whitespace from string columns
                df[col] = df[col].str.strip()
        
        logger.info("DataFrame cleaned successfully")
        return df
    
    @staticmethod
    def get_date_columns(df: pd.DataFrame) -> list:
        """
        Identify datetime columns in DataFrame.
        
        Args:
            df (pd.DataFrame): Input DataFrame
        
        Returns:
            list: List of column names with datetime dtype
        """
        return [col for col in df.select_dtypes(include=[np.datetime64]).columns]
    
    @staticmethod
    def validate_column_names(df: pd.DataFrame, validate_special_chars: bool = True,
                              max_length: int = 100) -> Dict[str, list]:
        """
        Validate column names for issues.
        
        Args:
            df (pd.DataFrame): Input DataFrame
            validate_special_chars (bool): Check for special characters
            max_length (int): Maximum allowed column name length
        
        Returns:
            Dict[str, list]: Dictionary of validation issues
        """
        issues = {
            'special_chars': [],
            'too_long': [],
            'duplicates': []
        }
        
        # Check for special characters
        if validate_special_chars:
            for col in df.columns:
                if re.search(r"\W", col):
                    issues['special_chars'].append(col)
        
        # Check for long names
        for col in df.columns:
            if len(col) > max_length:
                issues['too_long'].append(col)
        
        # Check for duplicates
        duplicates = df.columns[df.columns.duplicated()].tolist()
        if duplicates:
            issues['duplicates'] = list(set(duplicates))
        
        if any(issues.values()):
            logger.warning(f"Column name validation issues found: {issues}")
        
        return issues
    
    @staticmethod
    def detect_mixed_datatypes(df: pd.DataFrame) -> Dict[str, list]:
        """
        Detect columns with mixed datatypes.
        
        Args:
            df (pd.DataFrame): Input DataFrame
        
        Returns:
            Dict[str, list]: Dictionary of mixed datatype columns
        """
        mixed_cols = {'source': [], 'target': []}
        
        for col in df.columns:
            # Check if column has mixed types
            types = df[col].apply(type)
            if types.nunique() > 1:
                mixed_cols['source'].append(col)
        
        if any(mixed_cols.values()):
            logger.warning(f"Mixed datatype columns detected: {mixed_cols}")
        
        return mixed_cols
    
    @staticmethod
    def normalize_duplicate_column_names(df: pd.DataFrame) -> pd.DataFrame:
        """
        Remove ".1", ".2", etc. suffixes from duplicated column names.
        
        Args:
            df (pd.DataFrame): Input DataFrame
        
        Returns:
            pd.DataFrame: DataFrame with normalized column names
        """
        df = df.copy()
        df.columns = df.columns.str.split('.').str[0]
        logger.info("Column names normalized")
        return df
    
    @staticmethod
    def get_column_datatype_for_sql(df: pd.DataFrame, date_columns: list) -> str:
        """
        Generate SQL CREATE TABLE statement datatype string.
        
        Args:
            df (pd.DataFrame): Input DataFrame
            date_columns (list): List of column names with date datatype
        
        Returns:
            str: SQL datatype string for CREATE TABLE
        """
        sql_format = ""
        
        for col in df.columns:
            dt = df[col].dtype
            if dt == object:  # String datatype
                sql_format += f"{col} text, "
            elif col in date_columns:  # Date datatype
                sql_format += f"{col} date, "
            else:  # Numeric datatype
                sql_format += f"{col} real, "
        
        # Remove trailing comma and space
        sql_format = sql_format[:-2] if sql_format else ""
        
        logger.info("SQL datatype string generated")
        return sql_format
    
    @staticmethod
    def rename_first_column(df: pd.DataFrame, original_name: str) -> pd.DataFrame:
        """
        Rename first column to 'ColumnA' for SQL compatibility.
        
        Args:
            df (pd.DataFrame): Input DataFrame
            original_name (str): Store original column name
        
        Returns:
            pd.DataFrame: DataFrame with renamed first column
        """
        df = df.copy()
        if len(df.columns) > 0:
            df.rename(columns={df.columns[0]: "ColumnA"}, inplace=True)
            logger.info(f"First column renamed from '{original_name}' to 'ColumnA'")
        
        return df
