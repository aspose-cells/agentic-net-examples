// Title: Read a specific cell from an Excel file using zero‑based row and column indices with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# read cell value using row and column indexes | How to access a cell by zero based indices in Aspose.Cells .NET | Load password protected Excel workbook and retrieve cell data with Aspose.Cells C# | Get value from third row second column using Aspose.Cells worksheet.Cells | C# example for reading Excel cell with zero based indexing and password handling
// Tags: Aspose.Cells cell index access | C# extract Excel cell value | password protected workbook loading Aspose.Cells | worksheet.Cells row column usage | exception handling workbook load Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// This C# example demonstrates how to open an Excel workbook (including password‑protected files), select the first worksheet, access a cell using zero‑based row and column indexes, read its value, and write the result to the console, with proper error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Try to load the workbook without a password
            LoadOptions loadOptions = new LoadOptions();
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException)
        {
            // Assume the workbook is password‑protected; ask the user for a password
            Console.Write("The workbook is password protected. Enter password: ");
            string password = Console.ReadLine() ?? string.Empty;

            try
            {
                LoadOptions pwdOptions = new LoadOptions { Password = password };
                workbook = new Workbook(inputPath, pwdOptions);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Failed to open workbook with provided password: {innerEx.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Get the first worksheet (zero‑based index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define zero‑based row and column indices
            int rowIndex = 2;      // third row
            int columnIndex = 1;   // second column

            // Access the cell and read its value
            object cellValue = worksheet.Cells[rowIndex, columnIndex].Value;

            // Output the retrieved value
            Console.WriteLine($"Value at row {rowIndex}, column {columnIndex}: {cellValue}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing worksheet or cell: {ex.Message}");
        }
    }
}
