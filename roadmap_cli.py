"""
Entry point script for PyInstaller builds.
This script is used to create the executable.
"""
import multiprocessing

if __name__ == "__main__":
    """
    Entry point for PyInstaller builds.
    MUST be called at the very start for PyInstaller + multiprocessing on Windows
    Import here to avoid issues with subprocess re-imports
    """
    multiprocessing.freeze_support()
    from roadmap.main import run
    run()
