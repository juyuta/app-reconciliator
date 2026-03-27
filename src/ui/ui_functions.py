"""UI functions for window management and customization."""
import logging
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QColor

logger = logging.getLogger(__name__)

# Global state for window maximization
GLOBAL_STATE = 0


class UIFunctions:
    """Handles UI customization and window management functions."""
    
    @staticmethod
    def maximize_restore(main_window):
        """
        Toggle between maximized and normal window states.
        
        Args:
            main_window: MainWindow instance
        """
        global GLOBAL_STATE
        
        try:
            status = GLOBAL_STATE
            
            # IF NOT MAXIMIZED
            if status == 0:
                main_window.showMaximized()
                GLOBAL_STATE = 1
                
                # Remove margins when maximized
                try:
                    main_window.ui.body_layout.setContentsMargins(0, 0, 0, 0)
                    main_window.ui.maximize_button.setIcon(main_window.ui.icon9)
                except AttributeError:
                    logger.warning("UI elements not fully initialized")
            else:
                # Reset to normal state
                GLOBAL_STATE = 0
                main_window.showNormal()
                main_window.resize(main_window.width(), main_window.height())
                try:
                    main_window.ui.maximize_button.setIcon(main_window.ui.icon2)
                except AttributeError:
                    logger.warning("UI elements not fully initialized")
            
            logger.info(f"Window state changed. Maximized: {bool(GLOBAL_STATE)}")
            
        except Exception as e:
            logger.error(f"Error in maximize_restore: {str(e)}")
    
    @staticmethod
    def setup_ui_definitions(main_window):
        """
        Configure UI properties including frameless window, drop shadow, and button connections.
        
        Args:
            main_window: MainWindow instance
        """
        try:
            # Remove default title bar for custom styling
            main_window.setWindowFlag(QtCore.Qt.FramelessWindowHint)
            main_window.setAttribute(QtCore.Qt.WA_TranslucentBackground)
            
            # Create drop shadow effect
            main_window.shadow = QtWidgets.QGraphicsDropShadowEffect(main_window)
            main_window.shadow.setBlurRadius(20)
            main_window.shadow.setXOffset(0)
            main_window.shadow.setYOffset(0)
            main_window.shadow.setColor(QColor(0, 0, 0, 100))
            
            # Connect window control buttons
            if hasattr(main_window.ui, 'maximize_button'):
                main_window.ui.maximize_button.clicked.connect(
                    lambda: UIFunctions.maximize_restore(main_window)
                )
            
            if hasattr(main_window.ui, 'minimize_button'):
                main_window.ui.minimize_button.clicked.connect(
                    lambda: main_window.showMinimized()
                )
            
            if hasattr(main_window.ui, 'close_button'):
                main_window.ui.close_button.clicked.connect(
                    lambda: main_window.close()
                )
            
            logger.info("UI definitions configured successfully")
            
        except Exception as e:
            logger.error(f"Error setting up UI definitions: {str(e)}")
    
    @staticmethod
    def get_global_state():
        """
        Get current window state.
        
        Returns:
            int: 0 if normal, 1 if maximized
        """
        return GLOBAL_STATE
    
    @staticmethod
    def reset_global_state():
        """Reset global state to normal."""
        global GLOBAL_STATE
        GLOBAL_STATE = 0
        logger.info("Global state reset")
