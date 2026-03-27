"""Entry point for the Reconciliator application."""
import sys

from config.settings import setup_logging, get_logger
from config.constants import APP_NAME, APP_VERSION


def main():
    """Launch the Reconciliator GUI application."""
    # These imports are inside main() so the heavy GUI modules
    # are only loaded when actually running the app.
    from core.reconciliator import (
        global_exception_handler, show_error_dialog, MainWindow,
    )
    from PyQt5.QtWidgets import QApplication

    sys.excepthook = global_exception_handler

    setup_logging()
    logger = get_logger("main")
    logger.info(f"Starting {APP_NAME} v{APP_VERSION}")

    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        sys.exit(app.exec_())
    except Exception as e:
        logger.critical(f"Fatal application error: {e}", exc_info=True)
        show_error_dialog("Fatal Error", f"Application failed to start:\n{e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
