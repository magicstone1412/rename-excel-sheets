import openpyxl
import os
from tqdm import tqdm


def rename_sheets ( input_file_path, output_file_path, cell_reference='A1', num_chars=5, from_start=True ):
    """
    Rename Excel sheets based on characters from a cell's value and save to a new file.

    Parameters:
    - input_file_path: Path to the input Excel file
    - output_file_path: Path to the output Excel file
    - cell_reference: Cell containing the name (e.g., 'A1')
    - num_chars: Number of characters to take
    - from_start: True to take from start of string, False for end
    """
    try:
        # Check if input file exists
        if not os.path.exists ( input_file_path ):
            raise FileNotFoundError ( f"Input file '{input_file_path}' not found" )

        # Load the workbook
        workbook = openpyxl.load_workbook ( input_file_path )

        # Get total number of sheets for progress bar
        total_sheets = len ( workbook.worksheets )

        # Initialize progress bar
        print ( "Renaming sheets..." )
        progress_bar = tqdm ( total=total_sheets, desc="Processing", unit="sheet" )

        # Iterate through all sheets
        for sheet in workbook.worksheets:
            # Get the value from the specified cell
            cell_value = str ( sheet [ cell_reference ].value )

            if cell_value and cell_value != 'None':
                # Extract the specified number of characters
                if from_start:
                    new_name = cell_value [ :num_chars ]
                else:
                    new_name = cell_value [ -num_chars: ]

                # Clean the name (Excel sheet names have restrictions)
                new_name = ''.join ( c for c in new_name if c.isalnum () or c in (' ', '_') ).strip ()
                new_name = new_name [ :31 ]  # Excel sheet name max length is 31

                if new_name:  # Only rename if we have a valid name
                    sheet.title = new_name
                else:
                    print ( f"Warning: Empty or invalid name in {cell_reference} for sheet {sheet.title}" )
            else:
                print ( f"Warning: No value in {cell_reference} for sheet {sheet.title}" )

            # Update progress bar
            progress_bar.update ( 1 )

        # Close progress bar
        progress_bar.close ()

        # Save the modified workbook to a new file
        workbook.save ( output_file_path )
        print ( f"Excel file saved as '{output_file_path}' with new sheet names." )

    except Exception as e:
        print ( f"An error occurred: {str ( e )}" )
    finally:
        # Ensure progress bar is closed in case of error
        if 'progress_bar' in locals ():
            progress_bar.close ()


def main ():
    # Example configuration
    input_file = "2017.xlsx"  # Replace with your input Excel file path
    output_file = "output.xlsx"  # Replace with your desired output Excel file path
    cell_reference = 'A1'  # Cell containing the name
    num_chars = 9  # Number of characters to use
    from_start = False  # True for start, False for end

    # Call the rename function
    rename_sheets ( input_file, output_file, cell_reference, num_chars, from_start )


if __name__ == "__main__":
    main ()