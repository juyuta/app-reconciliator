"""File handling utilities for Reconciliator application."""
import os
from pathlib import Path
import logging
from typing import Tuple, Optional

logger = logging.getLogger(__name__)


class FileHandler:
    """Handles file operations for the Reconciliator application."""
    
    @staticmethod
    def get_file_name(file_path: str) -> str:
        """
        Extract filename from full file path.
        
        Args:
            file_path (str): Full file path
        
        Returns:
            str: Filename only
        """
        return os.path.basename(file_path)
    
    @staticmethod
    def get_file_extension(file_path: str) -> str:
        """
        Get file extension from path.
        
        Args:
            file_path (str): Full file path
        
        Returns:
            str: File extension (including dot)
        """
        return os.path.splitext(file_path)[1].lower()
    
    @staticmethod
    def is_valid_excel_file(file_path: str) -> bool:
        """
        Validate if file is a valid Excel file.
        
        Args:
            file_path (str): Full file path
        
        Returns:
            bool: True if valid Excel file, False otherwise
        """
        if not os.path.exists(file_path):
            logger.warning(f"File does not exist: {file_path}")
            return False
        
        extension = FileHandler.get_file_extension(file_path)
        valid_extensions = ('.xls', '.xlsx')
        
        if extension not in valid_extensions:
            logger.warning(f"Invalid file extension: {extension}")
            return False
        
        return True
    
    @staticmethod
    def ensure_directory_exists(directory: str) -> bool:
        """
        Ensure directory exists, create if needed.
        
        Args:
            directory (str): Directory path
        
        Returns:
            bool: True if successful
        """
        try:
            os.makedirs(directory, exist_ok=True)
            logger.info(f"Directory ensured: {directory}")
            return True
        except Exception as e:
            logger.error(f"Failed to create directory {directory}: {str(e)}")
            return False
    
    @staticmethod
    def delete_file(file_path: str) -> bool:
        """
        Safely delete a file.
        
        Args:
            file_path (str): Path to file to delete
        
        Returns:
            bool: True if successful
        """
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info(f"File deleted: {file_path}")
                return True
            else:
                logger.warning(f"File not found: {file_path}")
                return False
        except Exception as e:
            logger.error(f"Failed to delete file {file_path}: {str(e)}")
            return False
    
    @staticmethod
    def read_sql_file(file_path: str) -> Optional[str]:
        """
        Read SQL file contents.
        
        Args:
            file_path (str): Path to SQL file
        
        Returns:
            Optional[str]: File contents or None if error
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            logger.info(f"SQL file read successfully: {file_path}")
            return content
        except FileNotFoundError:
            logger.error(f"SQL file not found: {file_path}")
            return None
        except Exception as e:
            logger.error(f"Error reading SQL file {file_path}: {str(e)}")
            return None
    
    @staticmethod
    def get_directory_size(directory: str) -> int:
        """
        Calculate total size of directory in bytes.
        
        Args:
            directory (str): Directory path
        
        Returns:
            int: Total size in bytes
        """
        total_size = 0
        try:
            for dirpath, dirnames, filenames in os.walk(directory):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    if os.path.exists(filepath):
                        total_size += os.path.getsize(filepath)
        except Exception as e:
            logger.error(f"Error calculating directory size: {str(e)}")
        
        return total_size
