# Advance Analysis Tool

A financial advance payment analysis system for the Department of Homeland Security (DHS). This tool analyzes and validates advance payments across fiscal quarters, comparing current year (CY) and prior year (PY) data to ensure compliance and proper tracking.

## Features

- **Data Analysis**: Process and analyze DHS advance payment data from Excel files
- **Status Validation**: Validate advance payment status changes between periods
- **Compliance Checking**: Ensure compliance with DHS financial regulations
- **Period of Performance**: Track and validate anticipated liquidation dates
- **Abnormal Balance Detection**: Identify and flag unusual financial patterns
- **Multi-Component Support**: Analyze data for all DHS components (CBP, CG, CIS, etc.)
- **GUI Application**: User-friendly desktop interface built with tkinter
- **Excel Integration**: Seamless import/export with formatting preservation

## Installation

### Prerequisites

- Python 3.12 or higher
- Windows, macOS, or Linux operating system

### Install from Source

1. Clone the repository:
```bash
git clone https://github.com/yourusername/advance-analysis.git
cd advance-analysis
```

2. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install the package:
```bash
pip install -e .[all]
```

## Usage

### GUI Mode (Default)

Launch the main application window:
```bash
python -m advance_analysis
```

Or use the installed command:
```bash
advance-analysis
```

### Simplified GUI Mode

For trial balance sheet operations:
```bash
advance-analysis --simple
```

### Command Line Mode

Process files directly from the command line:
```bash
advance-analysis --cli \
    --input "path/to/advance_analysis.xlsx" \
    --component WMD \
    --quarter "FY25 Q2" \
    --current-tb "path/to/current_tb.xlsx" \
    --prior-tb "path/to/prior_tb.xlsx"
```

### Options

- `--log-level`: Set logging verbosity (DEBUG, INFO, WARNING, ERROR, CRITICAL)
- `--log-file`: Specify custom log file location
- `--no-log-file`: Disable file logging
- `--output`: Custom output directory (default: ~/Documents/Python Outputs)
- `--version`: Display version information

## Project Structure

```
advance-analysis/
├── src/
│   └── advance_analysis/
│       ├── __init__.py
│       ├── main.py              # Entry point
│       ├── core/                # Business logic
│       │   ├── cy_advance_analysis.py
│       │   ├── status_validations.py
│       │   └── data_processing.py
│       ├── modules/             # File handlers
│       │   ├── data_loader.py
│       │   ├── excel_handler.py
│       │   └── file_handler.py
│       ├── gui/                 # GUI components
│       │   ├── gui.py
│       │   ├── gui_window.py
│       │   └── gui_window_v2.py
│       └── utils/               # Utilities
│           ├── logging_config.py
│           └── theme_files.py
├── tests/                       # Test suite
├── docs/                        # Documentation
├── data/                        # Data files
├── scripts/                     # Helper scripts
├── assets/                      # Icons and resources
├── .github/                     # GitHub configuration
├── pyproject.toml              # Project configuration
├── .gitignore
├── CLAUDE.md                   # AI assistant instructions
└── README.md
```

## Development

### Setting Up Development Environment

1. Install development dependencies:
```bash
pip install -e .[dev]
```

2. Install pre-commit hooks:
```bash
pre-commit install
```

3. Run tests:
```bash
pytest
```

4. Run linting:
```bash
ruff check .
black --check .
mypy src/
```

### Code Style

This project follows:
- PEP 8 style guide
- Black formatting (100 character line limit)
- Type hints for all functions
- Google-style docstrings

### Testing

Run the test suite:
```bash
# All tests
pytest

# With coverage
pytest --cov=advance_analysis

# Specific test file
pytest tests/test_validations.py
```

## Configuration

The application stores configuration in `~/.advance_analysis_config.json`, including:
- Theme preferences
- Default paths
- Component settings

## Logging

Logs are stored in `~/Documents/Advance Analysis/logs/` with daily rotation.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

**Jéron Crooks**

## Acknowledgments

- Department of Homeland Security for the requirements and domain expertise
- The Python community for excellent libraries and tools

## Support

For issues, questions, or contributions, please visit the [GitHub repository](https://github.com/yourusername/advance-analysis).