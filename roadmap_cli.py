#!/usr/bin/env python
"""
Entry point script for PyInstaller builds.
This script is used to create the executable.
"""
import multiprocessing

if __name__ == "__main__":
    # MUST be called at the very start for PyInstaller + multiprocessing on Windows
    multiprocessing.freeze_support()
    
    # Import here to avoid issues with subprocess re-imports
    from roadmap.main import run
    run()
