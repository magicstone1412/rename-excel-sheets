# rename-excel-sheets
Rename Excel sheet names based on a referenced cell

## Pseudocode

```
FUNCTION rename_sheets(input_file_path, output_file_path, cell_reference, num_chars, from_start)
    TRY
        // Check if input file exists
        IF input_file_path DOES NOT EXIST
            RAISE ERROR "Input file not found"
        
        // Load the Excel workbook
        workbook ← LOAD_EXCEL_FILE(input_file_path)
        
        // Get total number of sheets
        total_sheets ← COUNT(workbook.sheets)
        
        // Initialize progress bar
        OUTPUT "Renaming sheets..."
        progress_bar ← CREATE_PROGRESS_BAR(total=total_sheets, description="Processing", unit="sheet")
        
        // Loop through each sheet in the workbook
        FOR EACH sheet IN workbook.sheets
            // Get the value from the specified cell and convert to string
            cell_value ← CONVERT_TO_STRING(sheet[cell_reference].value)
            
            // Check if cell value is not empty or null
            IF cell_value IS NOT EMPTY AND cell_value IS NOT "None"
                // Extract characters based on from_start
                IF from_start IS TRUE
                    new_name ← FIRST num_chars CHARACTERS OF cell_value
                ELSE
                    new_name ← LAST num_chars CHARACTERS OF cell_value
                
                // Clean the name for Excel sheet naming rules
                new_name ← KEEP ONLY alphanumeric, space, underscore IN new_name
                new_name ← TRIM whitespace FROM new_name
                new_name ← LIMIT new_name TO 31 CHARACTERS
                
                // Check if cleaned name is valid
                IF new_name IS NOT EMPTY
                    sheet.title ← new_name
                ELSE
                    OUTPUT "Warning: Empty or invalid name in cell_reference for sheet sheet.title"
            ELSE
                OUTPUT "Warning: No value in cell_reference for sheet sheet.title"
            
            // Update progress bar
            INCREMENT progress_bar BY 1
        
        // Close progress bar
        CLOSE progress_bar
        
        // Save the modified workbook to output file
        SAVE workbook TO output_file_path
        OUTPUT "Excel file saved as output_file_path with new sheet names"
        
    CATCH ANY error
        OUTPUT "Error occurred: error.message"
    FINALLY
        // Ensure progress bar is closed
        IF progress_bar EXISTS
            CLOSE progress_bar

FUNCTION main()
    // Define configuration
    input_file ← "input.xlsx"
    output_file ← "output.xlsx"
    cell_reference ← "A1"
    num_chars ← 5
    from_start ← TRUE
    
    // Call rename_sheets with configuration
    CALL rename_sheets(input_file, output_file, cell_reference, num_chars, from_start)

IF script is run directly
    CALL main()
```