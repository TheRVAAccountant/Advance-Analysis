"""
Main entry point for the Advance Analysis application.

This module provides the command-line interface and launches the GUI
for the Advance Analysis Tool.
"""

import sys
import argparse
import logging
from pathlib import Path
from typing import Optional

try:
    from .utils.logging_config import setup_logging, get_logger
    from .gui import run_gui, run_simplified_gui
    from .core.cy_advance_analysis import CYAdvanceAnalysis
except ImportError:
    # Fallback for direct script execution
    sys.path.insert(0, str(Path(__file__).parent.parent.parent))
    from src.advance_analysis.utils.logging_config import setup_logging, get_logger
    from src.advance_analysis.gui import run_gui, run_simplified_gui
    from src.advance_analysis.core.cy_advance_analysis import CYAdvanceAnalysis

logger = get_logger(__name__)


def parse_arguments() -> argparse.Namespace:
    """
    Parse command line arguments.
    
    Returns:
        Parsed command line arguments
    """
    parser = argparse.ArgumentParser(
        description="Advance Analysis Tool for DHS Financial Data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Launch the main GUI
  python -m advance_analysis
  
  # Launch the simplified GUI
  python -m advance_analysis --simple
  
  # Process files from command line
  python -m advance_analysis --cli --input file.xlsx --component WMD --quarter "FY25 Q2"
  
  # Set custom log level
  python -m advance_analysis --log-level DEBUG
"""
    )
    
    # GUI mode selection
    gui_group = parser.add_mutually_exclusive_group()
    gui_group.add_argument(
        "--simple",
        action="store_true",
        help="Launch the simplified GUI for trial balance operations"
    )
    gui_group.add_argument(
        "--cli",
        action="store_true",
        help="Run in command-line mode without GUI"
    )
    
    # CLI mode arguments
    cli_group = parser.add_argument_group("CLI Mode Options")
    cli_group.add_argument(
        "--input",
        type=str,
        help="Path to the advance analysis Excel file"
    )
    cli_group.add_argument(
        "--component",
        type=str,
        choices=["CBP", "CG", "CIS", "CYB", "FEM", "FLE", "ICE", "MGA", "MGT", "OIG", "TSA", "SS", "ST", "WMD"],
        help="DHS component code"
    )
    cli_group.add_argument(
        "--quarter",
        type=str,
        help="Fiscal year and quarter (e.g., 'FY25 Q2')"
    )
    cli_group.add_argument(
        "--current-tb",
        type=str,
        help="Path to current period DHSTIER trial balance"
    )
    cli_group.add_argument(
        "--prior-tb",
        type=str,
        help="Path to prior year DHSTIER trial balance"
    )
    cli_group.add_argument(
        "--output",
        type=str,
        help="Output directory path (defaults to project/outputs)"
    )
    
    # Logging options
    parser.add_argument(
        "--log-level",
        type=str,
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default="INFO",
        help="Set the logging level"
    )
    parser.add_argument(
        "--log-file",
        type=str,
        help="Custom log file path"
    )
    parser.add_argument(
        "--no-log-file",
        action="store_true",
        help="Disable file logging"
    )
    
    # Version
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 2.1.0"
    )
    
    return parser.parse_args()


def run_cli_mode(args: argparse.Namespace) -> int:
    """
    Run the application in CLI mode.
    
    Args:
        args: Parsed command line arguments
        
    Returns:
        Exit code (0 for success, 1 for error)
    """
    # Validate required arguments
    if not all([args.input, args.component, args.quarter]):
        logger.error("CLI mode requires --input, --component, and --quarter arguments")
        return 1
    
    try:
        logger.info(f"Processing advance analysis for {args.component} {args.quarter}")
        
        # Determine output directory
        if args.output:
            output_dir = Path(args.output)
        else:
            # Use project outputs directory
            project_root = Path(__file__).parent.parent.parent
            output_dir = project_root / "outputs"
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Create analyzer instance
        analyzer = CYAdvanceAnalysis(logger)
        
        # Note: The CYAdvanceAnalysis class would need to be updated to support
        # this CLI interface. For now, we'll just log a message.
        logger.warning("CLI mode is not fully implemented yet. Please use the GUI.")
        logger.info(f"Would process: {args.input} for {args.component} {args.quarter}")
        logger.info(f"Output would be saved to: {output_dir}")
        
        return 0
        
    except Exception as e:
        logger.error(f"Error processing data: {str(e)}", exc_info=True)
        return 1


def main() -> int:
    """
    Main entry point for the application.
    
    Returns:
        Exit code
    """
    args = parse_arguments()
    
    # Set up logging
    setup_logging(
        log_level=args.log_level,
        log_to_file=not args.no_log_file,
        log_filename=args.log_file
    )
    
    logger.info("Starting Advance Analysis Tool")
    
    try:
        if args.cli:
            # Run in CLI mode
            return run_cli_mode(args)
        elif args.simple:
            # Launch simplified GUI
            logger.info("Launching simplified GUI")
            run_simplified_gui()
        else:
            # Launch main GUI
            logger.info("Launching main GUI")
            run_gui()
        
        return 0
        
    except KeyboardInterrupt:
        logger.info("Application interrupted by user")
        return 130
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        return 1
    finally:
        logger.info("Advance Analysis Tool shutting down")


if __name__ == "__main__":
    sys.exit(main())