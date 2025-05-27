# Logs Directory

This directory contains application log files generated during runtime.

## Log File Format

Log files are named with the following pattern:
- `advance_analysis_YYYYMMDD_HHMMSS.log`

For example:
- `advance_analysis_20250527_143052.log` (created on May 27, 2025 at 14:30:52)

## Log Levels

The application logs at various levels:
- **DEBUG**: Detailed information for debugging
- **INFO**: General informational messages
- **WARNING**: Warning messages that don't prevent operation
- **ERROR**: Error messages for issues that need attention
- **CRITICAL**: Critical errors that may cause application failure

## Viewing Logs

You can view logs using any text editor or terminal commands:

```bash
# View the latest log
tail -f $(ls -t advance_analysis_*.log | head -1)

# Search for errors in all logs
grep ERROR advance_analysis_*.log

# View logs from today
ls advance_analysis_$(date +%Y%m%d)_*.log

# View the most recent log file
cat $(ls -t advance_analysis_*.log | head -1)
```

## Note

Log files in this directory ARE tracked by git. This allows for debugging and analysis of application behavior across different environments. Please ensure that no sensitive information (passwords, personal data, etc.) is logged in production environments.