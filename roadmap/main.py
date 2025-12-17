"""
CLI entry point for CE VHST Roadmap automation.

This module provides the command-line interface (CLI) for the roadmap management tool.
It parses command-line arguments and delegates operations to the RoadmapManager class.

The CLI supports the following commands:
    1. create - Create user interfaces
    2. delete - Delete interfaces (with optional archiving)
    3. cleanup - Remove interfaces for missing collaborators
    4. pointage - Export time tracking data
    5. update - Update conditional lists (LC)

The module integrates with Excel files using openpyxl, and can be called from both command-line and VBA macros.

Author: Mustapha EL KAMILI
"""
import sys

from roadmap.helpers import get_parser, logger
from roadmap.roadmap import RoadmapManager


def main() -> None:
    """
    Main entry point for CLI interface.

    Parses command-line arguments and executes the appropriate RoadmapManager operation.
    Supports 'create', 'delete', 'pointage', and 'update' commands.

    Default base directory is platform-specific:
        - Windows: 'C:\\Users\\MustaphaELKAMILI\\OneDrive - IKOSCONSULTING\\test_RM\\files'
        - Other: '/mnt/c/Users/MustaphaELKAMILI/OneDrive - IKOSCONSULTING/test_RM/files'

    Can be overridden with '--basedir' argument.

    Returns:
        None
    """
    logger.info("Roadmap Manager - Loading...")
    args = get_parser().parse_args()
    if args.basedir == 'none':
        logger.error("No base directory provided. Please use '--basedir' to specify the base directory.")
        sys.exit(1)
    else:
        manager = RoadmapManager(base_dir=args.basedir)

    if args.action == "create":
        if args.way == 'normal':
            manager.create_interfaces()
        elif args.way == 'para':
            manager.create_interfaces_fast()
        else:
            logger.error(f"Unknown '--way' argument '{args.way}'. Valid choices are 'normal' and 'para'.")
        return

    if args.action == "delete":
        if not args.force:
            logger.warning("⚠️  Operation not confirmed. Use --force to proceed.")
            return

        manager.delete_and_archive_interfaces(archive=args.archive)
        return

    if args.action == "cleanup":
        manager.delete_missing_collaborators()
        return

    if args.action == "pointage":
        manager.pointage()
        return

    if args.action == "update":
        try:
            manager.update_lc()
        except Exception as e:
            logger.error(f"Fatal error in update_lc: {e}", exc_info=True)
            sys.exit(1)

def run() -> None:
    """
    Entry point for console script installation.

    This function is called when the 'roadmap' command is invoked from the command line.
    It's registered as a console script in 'pyproject.toml'.

    Returns:
        None
    """
    main()


if __name__ == "__main__":
    main()
