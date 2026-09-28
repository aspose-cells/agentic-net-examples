// Title: Read a password‑protected ODS spreadsheet in read‑only mode with Aspose.Cells for .NET and extract non‑empty cell values
// AI Prompts: Generate C# code that creates a LoadOptions object with a password, opens an encrypted ODS workbook in read‑only mode using Aspose.Cells, and prints each populated cell. | Provide an example that catches the specific CellsException thrown for an incorrect password when opening a protected ODS file with Aspose.Cells. | Show how to traverse all worksheets of a password‑protected ODS spreadsheet and retrieve the address and string value of every non‑null cell using Aspose.Cells.
// Common Searches: asp.net open encrypted ods workbook using aspose.cells password | c# read only access to password protected ods spreadsheet Aspose.Cells | catch incorrect password exception when loading ODS file with Aspose.Cells | extract populated cells from protected ODS file in .NET
// Tags: load encrypted ODS with Aspose.Cells LoadOptions | verify ODS password Aspose.Cells .NET | read‑only ODS workbook extraction C# | iterate used range cells Aspose.Cells | handle CellsException password error

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the encrypted ODS file exists, sets the password via LoadOptions, loads the workbook in read‑only mode with Aspose.Cells, catches password‑related CellsException, then iterates each worksheet's used range and outputs the address and string value of every non‑empty cell.
class Program
{
    static void Main()
    {
        // Path to the encrypted ODS file
        string filePath = "encrypted_file.ods";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Password to open the file
        string password = "myPassword";

        // Load options: specify ODS format and password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
        {
            Password = password
        };

        Workbook workbook;
        try
        {
            // Load the workbook with the provided password
            workbook = new Workbook(filePath, loadOptions);
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for password errors and other issues
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through each worksheet and extract cell values
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Determine the used range of the worksheet
                var usedRange = sheet.Cells.MaxDisplayRange; // Use var to avoid ambiguity with System.Range
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Skip empty cells
                        if (cell.Type != CellValueType.IsNull)
                        {
                            // Output cell address and its string representation
                            Console.WriteLine($"Cell {cell.Name}: {cell.StringValue}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
