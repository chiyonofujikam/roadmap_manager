"""
    Roadmap Management Tool for CE VHST.

    A Python package for automating roadmap management tasks including time
    tracking export, interface creation, and data synchronization.

    Main Components:
        - RoadmapManager: Core class for managing roadmap operations
        - main: CLI entry point function

    Example:
        >>> from roadmap import RoadmapManager
        >>> manager = RoadmapManager(base_dir="/path/to/files")
        >>> manager.create_interfaces(archive=True)

Author: Mustapha EL KAMILI
"""

__version__ = "1.0.0"

from roadmap.roadmap import RoadmapManager
from roadmap.main import main

__all__ = ["RoadmapManager", "main"]
