"""Custom widgets for the Reconciliator application."""
import logging
from PyQt5 import QtCore
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtWidgets import QWidget, QLabel

from config.constants import ICON_PATHS

logger = logging.getLogger(__name__)


class LoadingScreen(QWidget):
    """
    Loading animation widget displayed during long-running operations.
    Shows a GIF animation to provide visual feedback to the user.
    """
    
    def __init__(self, parent=None, icon_path: str = None):
        """
        Initialize loading screen.
        
        Args:
            parent: Parent widget
            icon_path (str): Path to loading GIF icon
        """
        super().__init__(parent)
        
        if icon_path is None:
            icon_path = ICON_PATHS["loading"]
        
        try:
            # Set window properties
            self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
            self.setWindowFlags(
                QtCore.Qt.FramelessWindowHint |
                QtCore.Qt.Window |
                QtCore.Qt.WindowStaysOnTopHint
            )
            self.setFixedSize(200, 200)
            
            # Create label and animation
            label = QLabel(self)
            self.movie = QMovie(icon_path)
            label.setMovie(self.movie)
            label.setAlignment(Qt.AlignCenter)
            
            # Start animation
            self.movie.start()
            self.activateWindow()
            self.show()
            
            logger.info("Loading screen created successfully")
            
        except Exception as e:
            logger.error(f"Error creating loading screen: {str(e)}")
            raise
    
    def close_loading(self):
        """Close the loading screen."""
        if self.movie:
            self.movie.stop()
        self.close()
        logger.info("Loading screen closed")
