# Logs Directory

This directory contains application log files generated during runtime.

## Log File Format

Log files are named with the following pattern:
- `advance_analysis_YYYYMMDD.log`

For example:
- `advance_analysis_20250527.log`

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
tail -f advance_analysis_*.log

# Search for errors
grep ERROR advance_analysis_*.log

# View today's log
cat advance_analysis_$(date +%Y%m%d).log
```

## Note

Log files in this directory are NOT tracked by git (only this README and .gitkeep are tracked). This ensures that potentially sensitive operational data in logs remains local to each installation.