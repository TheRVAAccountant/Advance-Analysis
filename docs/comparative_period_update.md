# Comparative Period Logic Update

## Overview
Updated the comparative period calculation logic in the Advance Analysis application to align with new business requirements.

## File Modified
- **Location**: `src/advance_analysis/core/data_processing.py`
- **Function**: `get_comparative_period()`

## Logic Changes

### Previous Logic
- Q1 → Prior Year Q4
- Q2 → Prior Year Q4
- Q3 → Current Year Q2
- Q4 → Current Year Q3

### New Logic
- **Q1 → Prior Year Q3** (e.g., FY25 Q1 compares to FY24 Q3)
- **Q2 → Prior Year Q3** (e.g., FY25 Q2 compares to FY24 Q3)
- **Q3 → Current Year Q2** (e.g., FY25 Q3 compares to FY25 Q2)
- **Q4 → Current Year Q3** (e.g., FY25 Q4 compares to FY25 Q3)

## Rationale
The updated logic ensures that:
- First and second quarter comparisons now reference Q3 of the prior fiscal year instead of Q4
- Third and fourth quarter comparisons continue to reference the previous quarter within the same fiscal year
- This provides a more consistent comparative baseline for advance payment analysis

## Impact
This change affects:
1. Which prior period data is used for comparison in advance payment validations
2. The selection of comparative DHSTIER trial balance files
3. All status validation comparisons between current and prior periods

## Testing Recommendations
When testing this change:
1. Verify Q1 and Q2 analyses correctly reference prior year Q3 data
2. Confirm Q3 and Q4 analyses correctly reference current year previous quarter
3. Check that all validation rules properly use the updated comparative periods