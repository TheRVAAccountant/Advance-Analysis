# How to Run the Advance Analysis Application

## Method 1: Using the Run Script (Recommended)

From the project root directory, run:

```bash
python run_advance_analysis.py
```

or

```bash
python3 run_advance_analysis.py
```

## Method 2: Running as a Python Module

From the project root directory, run:

```bash
python -m src.advance_analysis.main
```

or

```bash
python3 -m src.advance_analysis.main
```

## Method 3: Using the GUI Module Directly

From the project root directory, run:

```bash
python -m src.advance_analysis.gui.run_gui
```

## Common Issues

1. **Import Errors**: Make sure you're running from the project root directory, not from within the src folder.

2. **Module Not Found**: Clean Python cache files:
   ```bash
   find . -type d -name __pycache__ -exec rm -rf {} +
   ```

3. **Windows COM Warning**: This warning is normal on non-Windows systems and can be ignored:
   ```
   Windows COM modules not available - Excel COM automation features will be disabled
   ```

## Command Line Options

- `--simple` - Launch the simplified GUI
- `--log-level DEBUG` - Enable debug logging
- `--help` - Show all available options