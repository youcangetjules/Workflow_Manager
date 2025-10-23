# LATA Lookup Troubleshooting Guide

## Issue: LATA Not Autopopulating from CLLI Selection

### **Problem Description**
When selecting a CLLI code, the City field populates but the LATA field remains empty or doesn't populate correctly.

## **Enhanced LATA Lookup**

### **✅ Improved Column Detection**
The system now tries multiple approaches to find LATA data:

1. **Keyword-based Search**: Looks for columns containing:
   - `lata` (case insensitive)
   - `area` (case insensitive)
   - `zone` (case insensitive)
   - `region` (case insensitive)

2. **Numeric Column Detection**: If no LATA column found with keywords:
   - Searches all remaining columns
   - Looks for numeric or alphanumeric values
   - Identifies potential LATA codes (≤10 characters with digits)

3. **Data Validation**: Ensures data is not NaN, None, or empty

### **✅ Enhanced Debug Output**
The system now provides detailed debugging information:

```
Attempting to autopopulate from CLLI: ABC12345
Excel data columns: ['Host CLLI', 'City', 'LATA', 'State', 'Region']
Found 1 matching rows
Row data: {'Host CLLI': 'ABC12345', 'City': 'Los Angeles', 'LATA': '730', 'State': 'CA', 'Region': 'West'}
Potential City columns: ['City']
Found City column 'City' with value: 'Los Angeles'
Potential LATA columns: ['LATA']
Found LATA column 'LATA' with value: '730'
Final autopopulation result:
  CLLI: ABC12345
  City: 'Los Angeles'
  LATA: '730'
```

## **Troubleshooting Steps**

### **Step 1: Check Excel File Structure**
Run the application and check console output for:
```
Loaded X rows from Excel file
Excel columns: ['Column1', 'Column2', ...]
First few rows:
   Column1  Column2  Column3
0  Value1   Value2   Value3
```

### **Step 2: Check Column Names**
Look for columns that might contain LATA data:
- **Direct**: `LATA`, `Lata`, `lata`
- **Related**: `Area`, `Zone`, `Region`, `Code`
- **Numeric**: Any column with numeric values

### **Step 3: Check Data Format**
Verify LATA data format in Excel:
- **Numeric**: `730`, `1234`
- **Text**: `LATA730`, `Area-123`
- **Mixed**: `730A`, `123-B`

### **Step 4: Test CLLI Selection**
1. Select a CLLI from dropdown
2. Check console output for debug information
3. Verify if LATA column is found
4. Check if LATA value is populated

## **Common Issues and Solutions**

### **Issue 1: No LATA Column Found**
**Symptoms**:
- City populates but LATA remains empty
- Console shows "No LATA found with keywords"

**Solutions**:
- Check Excel column names
- Rename column to include "LATA" or "Area"
- Verify data is not empty/NaN

### **Issue 2: LATA Column Found but Empty**
**Symptoms**:
- Console shows LATA column found
- But LATA value is empty or "nan"

**Solutions**:
- Check Excel data for empty cells
- Fill in missing LATA data
- Verify data format

### **Issue 3: Wrong LATA Column Selected**
**Symptoms**:
- LATA populates but with wrong data
- Multiple columns match LATA keywords

**Solutions**:
- Rename columns to be more specific
- Check data in Excel file
- Verify correct LATA column

### **Issue 4: Data Format Issues**
**Symptoms**:
- LATA column found but value not populated
- Console shows data but field remains empty

**Solutions**:
- Check for special characters
- Verify data type in Excel
- Clean data format

## **Excel File Requirements**

### **Recommended Column Names**
For best results, use these column names:
- **CLLI**: `Host CLLI`, `CLLI`, `CLLI Code`
- **City**: `City`, `Location`, `Place`
- **LATA**: `LATA`, `Area`, `Zone`, `Region`

### **Data Format**
- **LATA Values**: Should be numeric or short alphanumeric
- **No Empty Cells**: Ensure LATA data is populated
- **Consistent Format**: Use same format for all LATA values

### **Example Excel Structure**
```
Host CLLI    City        LATA    State
ABC12345     Los Angeles 730     CA
XYZ67890     Houston     720     TX
DEF45678     Chicago     312     IL
```

## **Testing the Fix**

### **Test 1: Basic LATA Lookup**
1. Select a CLLI from dropdown
2. Check console for debug output
3. Verify LATA field populates
4. Check if value matches Excel data

### **Test 2: Multiple LATA Columns**
1. If Excel has multiple LATA-related columns
2. Check which column is selected
3. Verify correct LATA value is used
4. Adjust column names if needed

### **Test 3: Data Format Validation**
1. Test with different LATA formats
2. Check numeric vs text values
3. Verify special characters are handled
4. Test with empty/null values

## **Debug Information**

### **Console Output to Check**
```
Attempting to autopopulate from CLLI: [CLLI_CODE]
Excel data columns: [LIST_OF_COLUMNS]
Found X matching rows
Row data: {DICT_OF_ROW_DATA}
Potential City columns: [CITY_COLUMNS]
Found City column 'COLUMN_NAME' with value: 'CITY_VALUE'
Potential LATA columns: [LATA_COLUMNS]
Found LATA column 'COLUMN_NAME' with value: 'LATA_VALUE'
Final autopopulation result:
  CLLI: [CLLI_CODE]
  City: '[CITY_VALUE]'
  LATA: '[LATA_VALUE]'
```

### **What to Look For**
1. **Column Detection**: Are LATA columns found?
2. **Data Values**: Are LATA values populated?
3. **Field Updates**: Do City and LATA fields update?
4. **Error Messages**: Any exceptions or errors?

## **Manual Testing**

### **Test with Specific CLLI**
1. Find a CLLI code in Excel file
2. Note the expected City and LATA values
3. Select that CLLI in the application
4. Verify both fields populate correctly

### **Test with Different Data**
1. Try multiple CLLI codes
2. Check if LATA populates consistently
3. Verify data matches Excel file
4. Test edge cases (empty values, special characters)

## **If LATA Still Not Working**

### **Check Excel File**
1. Open Excel file manually
2. Verify LATA column exists
3. Check data format and values
4. Ensure no empty cells

### **Check Console Output**
1. Run application
2. Select a CLLI
3. Check console for debug information
4. Look for error messages

### **Verify Column Names**
1. Check exact column names in Excel
2. Compare with console output
3. Rename columns if needed
4. Test again

The enhanced LATA lookup should now work with various column names and data formats. Check the console output for detailed debugging information!
