# CLAUDE.md - AI Assistant Guide for Finansbot

> **Last Updated**: 2025-11-17
> **Project**: Finansbot - Financial Consolidation Report Update Tool
> **Language**: Python 3.12+
> **Package Manager**: UV

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Codebase Structure](#codebase-structure)
3. [Core Business Logic](#core-business-logic)
4. [Development Workflows](#development-workflows)
5. [Key Conventions](#key-conventions)
6. [Common Tasks](#common-tasks)
7. [Testing & Validation](#testing--validation)
8. [Troubleshooting](#troubleshooting)

---

## Project Overview

### Purpose
Finansbot is a specialized financial automation tool designed to update consolidation reports for Turkish holding companies. It processes monthly financial data and performs intelligent data transformations based on ownership percentages.

### Key Functionality
- **Data Transformation**: Converts 40% ownership data to 100% consolidated figures (multiply by 2.5)
- **Multi-Sheet Updates**: Updates both raw data sheets and formula-based reporting sheets
- **Intelligent Month Detection**: Automatically identifies the latest available data month
- **Dual Interface**: Command-line tool and web-based Streamlit interface

### Business Context
The tool handles 8 companies with different reporting methodologies:
- **3 companies** report 100% data directly (no transformation needed)
- **5 companies** report 40% data (requires 2.5x multiplication for consolidation)

---

## Codebase Structure

```
finansbot/
├── update_konsolidasyon.py    # Core CLI application (327 lines)
│   ├── Business logic for data reading
│   ├── Excel sheet update functions
│   └── Formula update algorithms
│
├── app/
│   └── streamlit_app.py       # Web UI (98 lines)
│       ├── File upload handling
│       ├── Temporary file management
│       └── Download functionality
│
├── pyproject.toml             # UV project configuration
├── uv.lock                    # Dependency lock file
├── README.md                  # User documentation (Turkish)
└── CLAUDE.md                  # This file - AI assistant guide
```

### Key Files Deep Dive

#### `update_konsolidasyon.py` (Core Module)
**Critical Constants:**
- `COMPANY_MAPPING` (lines 26-35): Maps source rows to target rows with transformation flags
- `MONTH_MAPPING` (lines 39-52): Maps data columns to consolidation columns (12 months)

**Key Functions:**
- `find_last_month_with_data(ws)` (lines 55-74): Detects latest month with actual data
- `read_data_from_data_xlsx(data_file)` (lines 77-119): Reads and transforms source data
- `update_gercaylık_euro_sheet(all_data, konsolidasyon_file)` (lines 122-161): Updates raw data sheet
- `update_finansal_ay_formulas(last_month_name, konsolidasyon_file)` (lines 163-252): Updates formula-based reporting sheet
- `run_update(data_file, konsolidasyon_file, output_file)` (lines 254-274): Main orchestration function

#### `app/streamlit_app.py` (Web Interface)
**Key Features:**
- Temporary file handling for uploads (lines 27-42)
- Non-destructive processing (creates new output files)
- Cleanup mechanisms for temporary resources

---

## Core Business Logic

### Data Transformation Rules

#### Company Classification
```python
# 100% Data Companies (no transformation)
- İnci Holding      (Row 5 in data.xlsx → Row 5 in consolidation)
- ISM               (Row 6 → Row 6)
- İncitaş           (Row 7 → Row 7)

# 40% Data Companies (multiply by 2.5)
- İnci GS Yuasa              (Row 8 → Row 8)
- Yusen İnci Lojistik        (Row 9 → Row 9)
- Maxion Jantas              (Row 10 → Row 10)
- Maxion Celik               (Row 11 → Row 11)
- Maxion Aluminyum           (Row 12 → Row 12)
```

#### Month Column Mapping
Data file uses **odd columns** (5, 7, 9...) for "Gerçekleşen" (Actual) values:
- Column 5 (Ocak/January) → Target Column B (2)
- Column 7 (Şubat/February) → Target Column C (3)
- ...continuing through December

### Excel Sheet Operations

#### Sheet 1: "gerç aylık-eur" (Actual Monthly - EUR)
- **Purpose**: Store raw monthly actual values
- **Update Strategy**: Direct cell value updates (no formulas)
- **Range**: Rows 5-12, Columns B-M (12 months)

#### Sheet 2: "Finansal Raporlama AY" (Monthly Financial Reporting)
- **Purpose**: Formula-based analysis comparing budget vs actual
- **Update Strategy**: Intelligent formula rewriting
- **Complex Logic**: Updates column references in formulas to point to latest month
- **Formula Columns**: D, E, G, H, J, K, M, N (alternating metrics and comparisons)

### Formula Update Algorithm (Lines 163-252)

**Critical Understanding**: The formula update system is particularly sophisticated:

1. **Month-to-Column Mapping**: Maps Turkish month names to Excel column pairs (Budget, Actual)
   ```python
   "Ekim": ("AC", "AD")  # October → Columns AC (Budget), AD (Actual)
   ```

2. **Two-Phase Replacement Strategy**:
   - Phase 1: Replace all month column references with temporary placeholders
   - Phase 2: Replace placeholders with target month columns

3. **Why This Approach?**: Prevents accidental overwrites when column names overlap (e.g., 'A' in 'AA')

**Example**: If latest month is October (Ekim):
```python
# Original formula referencing September (Z, AA):
='A -Döviz'!AA6

# Updated formula referencing October (AC, AD):
='A -Döviz'!AD6
```

---

## Development Workflows

### Setting Up Development Environment

```bash
# Clone repository
git clone <repository-url>
cd finansbot

# Install dependencies using UV
uv sync

# Verify installation
uv run python -c "import openpyxl, streamlit; print('Dependencies OK')"
```

### Running the Application

#### CLI Mode
```bash
# Basic usage (uses default files in project root)
uv run update_konsolidasyon.py

# With custom files
uv run update_konsolidasyon.py \
  --data /path/to/data.xlsx \
  --konsolidasyon /path/to/konsolidasyon.xlsx \
  --output /path/to/output.xlsx
```

#### Web Interface Mode
```bash
# Start Streamlit server
uv run streamlit run app/streamlit_app.py

# Access at http://localhost:8501
```

### Git Workflow

**Important Branch Conventions**:
- Development branches follow pattern: `claude/claude-md-<session-id>`
- Current branch: `claude/claude-md-mi339k4f57dif9w5-01DSqwTFUUjAFreuMkJdzpgw`
- Always push with: `git push -u origin <branch-name>`

**Commit Message Style** (based on git log analysis):
- Use descriptive, action-oriented messages
- Focus on user-facing changes
- Example: "Refine Streamlit app for updating consolidation reports: updated README for clarity..."

---

## Key Conventions

### Code Style

1. **Turkish Naming in Business Logic**
   - Sheet names, month names, and company names use Turkish
   - This is intentional - matches source Excel files
   - Keep Turkish identifiers unchanged unless explicitly refactoring

2. **Function Naming**
   - Use descriptive English names for functions
   - Use snake_case consistently
   - Example: `update_gercaylık_euro_sheet` (mixed for clarity)

3. **Error Handling**
   - Use try-except in main() with traceback printing
   - Provide clear Turkish error messages for end users
   - Exit with status code 1 on errors

4. **Print Statements for User Feedback**
   - Heavy use of print statements with visual separators (===)
   - Progress indicators during long operations
   - Before/after value comparisons for transparency

### Data Handling Conventions

1. **Always Use `data_only=True`** when loading workbooks
   ```python
   wb = load_workbook(data_file, data_only=True)
   ```
   This ensures formulas are evaluated to values.

2. **Cell Value Checking Pattern**
   ```python
   value = ws.cell(row, col).value
   if value and value != 0:  # Check both None and zero
       # Process value
   ```

3. **Column Reference Conversions**
   ```python
   from openpyxl.utils import get_column_letter
   col_letter = get_column_letter(col_idx)  # 1 → 'A', 27 → 'AA'
   ```

### Excel File Expectations

**Input File: `data.xlsx`**
- Sheet name: "Export"
- Row 5-12: Company data
- Columns 5, 7, 9, 11... (odd numbers): "Gerçekleşen" (Actual) monthly values
- Columns 4, 6, 8, 10... (even numbers): "Bütçe" (Budget) - ignored by script

**Output File: Consolidation Report**
- Sheet 1: "gerç aylık-eur" (raw data, rows 5-12, cols B-M)
- Sheet 2: "Finansal Raporlama AY" (formulas, rows 6-14, specific columns)
- Sheet 3: "A -Döviz" (referenced by formulas, not modified)

---

## Common Tasks

### Task 1: Adding a New Company

1. **Update COMPANY_MAPPING** (update_konsolidasyon.py:26)
   ```python
   COMPANY_MAPPING = {
       # ... existing companies ...
       13: (13, "New Company Name", True),  # True if 40% data
   }
   ```

2. **Consider**: Do formulas in "Finansal Raporlama AY" need updates?

3. **Test**: Verify both sheets update correctly

### Task 2: Adding a New Month to Year

The system handles 12 months automatically. For a new fiscal year:

1. **Update MONTH_MAPPING** if column positions change
2. **Update month_to_columns** in `update_finansal_ay_formulas()` (line 170)
3. **Verify** formula column mappings in target workbook

### Task 3: Modifying Transformation Logic

Current: 40% → 100% uses multiplier 2.5

To change:
1. Locate line 102: `converted_value = value * 2.5`
2. Modify multiplier or add conditional logic
3. Update README.md with new business rules
4. Test with real data files

### Task 4: Adding New Metrics/Sheets

1. Create new update function following pattern:
   ```python
   def update_new_sheet(data, konsolidasyon_file):
       print("\n" + "=" * 80)
       print("NEW SHEET UPDATING...")
       print("=" * 80)

       wb = load_workbook(konsolidasyon_file)
       ws = wb['SheetName']

       # Update logic here

       wb.save(konsolidasyon_file)
       wb.close()
   ```

2. Add call to `run_update()` function (line 254)

3. Update README.md feature list

### Task 5: Debugging Excel Formula Issues

**Diagnosis Steps**:

1. **Check Formula Before/After**
   ```python
   # Add debugging print in update_finansal_ay_formulas()
   print(f"Old: {old_formula}")
   print(f"New: {new_formula}")
   ```

2. **Verify Column Mappings**
   - Open target Excel file manually
   - Check actual column letters in "A -Döviz" sheet
   - Compare with `month_to_columns` dict (line 170)

3. **Test Formula in Excel**
   - Copy updated formula
   - Paste in Excel to verify it evaluates correctly

---

## Testing & Validation

### Manual Testing Checklist

Since this project lacks automated tests, use this checklist:

- [ ] **Data Reading**: Verify all 8 companies read correctly from data.xlsx
- [ ] **Month Detection**: Confirm correct latest month identified
- [ ] **40% Conversion**: Check 5 companies have values multiplied by 2.5
- [ ] **100% Pass-Through**: Check 3 companies unchanged
- [ ] **Sheet 1 Update**: Open output file, verify "gerç aylık-eur" values
- [ ] **Sheet 2 Formulas**: Open output file, click cells in "Finansal Raporlama AY", verify formulas point to correct columns
- [ ] **All Months Updated**: Check not just latest month, but all months with data

### Test Data Validation

**Before Running Script**:
```bash
# Check data file exists and has expected structure
uv run python -c "
from openpyxl import load_workbook
wb = load_workbook('data.xlsx', data_only=True)
ws = wb['Export']
print(f'Companies: {[ws.cell(i, 1).value for i in range(5, 13)]}')
print(f'Last month data in row 5: {ws.cell(5, 21).value}')  # Sept
"
```

**After Running Script**:
```bash
# Verify output file created and has updates
uv run python -c "
from openpyxl import load_workbook
wb = load_workbook('Konsolidasyon_2025_NV (1)_guncel.xlsx')
ws = wb['gerç aylık-eur']
print(f'Company in row 5: {ws.cell(5, 1).value}')
print(f'September value: {ws.cell(5, 10).value}')  # Column J
"
```

### Known Edge Cases

1. **Empty Data Months**: Script correctly stops processing when encountering first empty month
2. **Missing Sheets**: Will crash with KeyError - ensure target file has correct sheet names
3. **Formula Edge Cases**: Complex formulas with nested parentheses may not update correctly
4. **File Locks**: Ensure Excel files are closed before running script

---

## Troubleshooting

### Common Errors

#### Error: "ModuleNotFoundError: No module named 'openpyxl'"
```bash
# Solution: Reinstall dependencies
uv sync
```

#### Error: "❌ HATA: data.xlsx bulunamadı!"
```bash
# Solution: Ensure running from project root or use --data flag
uv run update_konsolidasyon.py --data /full/path/to/data.xlsx
```

#### Error: KeyError accessing sheet
```python
# Cause: Sheet name mismatch in target Excel file
# Solution: Verify sheet names match:
# - "gerç aylık-eur" (note: lowercase, Turkish characters)
# - "Finansal Raporlama AY"
# - "A -Döviz"
```

#### Error: Permission denied writing file
```bash
# Cause: Excel file open in another application
# Solution: Close Excel application, retry
```

### Performance Considerations

- **File Size**: Typical Excel files (150KB) process in < 5 seconds
- **Memory**: Openpyxl loads entire workbook into memory (~50MB per file)
- **Optimization Tip**: For very large files, consider openpyxl's read_only/write_only modes

### Debugging Tips

1. **Add Verbose Logging**:
   ```python
   # Insert at top of functions
   import logging
   logging.basicConfig(level=logging.DEBUG)
   ```

2. **Inspect Cell Values**:
   ```python
   # Add temporary debugging code
   print(f"Cell {row},{col}: {ws.cell(row, col).value} (type: {ws.cell(row, col).data_type})")
   ```

3. **Test with Copy**:
   - Always test with copies of production files
   - Use --output flag to avoid overwriting originals

---

## Dependencies

### Production Dependencies

```toml
[project.dependencies]
openpyxl = ">=3.1.5"     # Excel file manipulation
streamlit = ">=1.32.0"   # Web interface
```

### Why These Versions?

- **openpyxl 3.1.5+**: Required for proper formula handling in newer Excel formats
- **streamlit 1.32.0+**: File upload stability improvements

### Adding New Dependencies

```bash
# Add to pyproject.toml dependencies list, then:
uv sync

# To add specific version:
uv add "package_name>=version"
```

---

## AI Assistant Best Practices

### When Modifying This Codebase:

1. **Understand Business Logic First**
   - This is financial data - errors have real consequences
   - Always verify the 40% vs 100% company classification
   - Test calculations manually before committing

2. **Preserve Turkish Terminology**
   - Don't translate sheet names, company names, or month names
   - These must match the Excel files exactly
   - Comments can be in English, but identifiers stay in Turkish

3. **Maintain Backward Compatibility**
   - Users have existing Excel files with specific structures
   - Changes to column mappings or sheet names break existing workflows
   - If structure must change, create migration path

4. **Test Both Interfaces**
   - Changes to core logic affect both CLI and Streamlit
   - Verify both entry points after modifications

5. **Document Assumptions**
   - This code assumes specific Excel file structures
   - Document any new assumptions in code comments
   - Update this CLAUDE.md file with significant changes

6. **Handle Errors Gracefully**
   - Users are finance professionals, not developers
   - Error messages must be clear and actionable
   - Provide Turkish translations for user-facing errors

### Code Review Checklist

Before committing changes:

- [ ] Code follows existing snake_case convention
- [ ] Turkish identifiers preserved where needed
- [ ] Print statements added for user feedback on long operations
- [ ] Error handling includes traceback for debugging
- [ ] Both CLI and Streamlit interfaces tested
- [ ] README.md updated if user-visible changes
- [ ] This CLAUDE.md updated if architecture changes
- [ ] Manual testing completed with real Excel files
- [ ] Git commit message is descriptive and action-oriented

---

## Future Enhancements (Potential)

Based on codebase analysis, potential improvements:

1. **Automated Testing**
   - Unit tests for data transformation logic
   - Integration tests with sample Excel files
   - Consider: pytest framework

2. **Configuration File**
   - Move COMPANY_MAPPING and MONTH_MAPPING to YAML/JSON
   - Allow non-developers to modify company lists
   - Consider: config.yaml in project root

3. **Logging System**
   - Replace print statements with proper logging
   - Configurable log levels (INFO, DEBUG, ERROR)
   - Log file output option for troubleshooting

4. **Data Validation**
   - Pre-flight checks on Excel file structure
   - Verify expected sheets exist before processing
   - Validate company names match expected list

5. **Progress Indicators**
   - Streamlit progress bar for long operations
   - ETA calculations for large files
   - Consider: tqdm for CLI progress bars

6. **Backup Mechanism**
   - Automatic backup before overwriting files
   - Configurable backup retention (e.g., keep last 5)
   - Restore functionality for accidents

---

## Support & Maintenance

### Project Status
- **Active Development**: Yes
- **Primary Use Case**: Internal financial reporting tool
- **User Base**: Finance team members (non-technical)

### Key Contacts
- **Technical Owner**: [Based on git history - check git log for contributors]
- **Business Owner**: Finance department using the tool

### Version History
- **v1.0.0** (Current): Initial release with dual interface (CLI + Streamlit)
- Based on git commits:
  - 71fa6e5: Streamlit refinements and README updates
  - 9848cf2: Added Streamlit interface
  - 06fd4fe: Initial commit

---

## Quick Reference

### Most Important Files
1. `update_konsolidasyon.py` - Core logic, start here for understanding
2. `README.md` - User documentation in Turkish
3. `app/streamlit_app.py` - Web interface wrapper

### Most Important Functions
1. `run_update()` - Main orchestration (line 254)
2. `update_gercaylık_euro_sheet()` - Raw data updates (line 122)
3. `update_finansal_ay_formulas()` - Formula updates (line 163)

### Most Important Constants
1. `COMPANY_MAPPING` - Company-to-row mapping with 40% flags (line 26)
2. `MONTH_MAPPING` - Month-to-column mapping (line 39)
3. `month_to_columns` - Formula column mapping (line 170)

### Critical Business Rules
- 40% companies multiply by 2.5 (line 102)
- Only "Gerçekleşen" (Actual) columns processed (odd column numbers)
- Updates ALL months, not just latest (lines 133-153)
- Formula updates target specific columns (D, E, G, H, J, K, M, N)

---

**End of CLAUDE.md**

*This document should be updated whenever significant architectural or business logic changes are made to the codebase.*
