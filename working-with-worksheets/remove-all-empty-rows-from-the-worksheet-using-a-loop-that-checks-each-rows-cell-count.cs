// Title: Remove blank rows from an Excel worksheet with Aspose.Cells for .NET using a reverse‑order loop
// AI Prompts: Write C# code that uses Aspose.Cells to scan a worksheet from the last row upward and delete any rows that contain only empty cells. | Provide a method that checks each cell in a row for null or empty values and removes the row via Aspose.Cells when no data is found. | Show how to safely delete multiple empty rows in an Excel file with Aspose.Cells without affecting row indices.
// Common Searches: C# Aspose.Cells how to delete rows that have no data in an Excel sheet | remove blank rows from .xlsx using Aspose.Cells library | loop through rows in reverse and delete empty rows Aspose.Cells .NET example | Aspose.Cells delete rows with only null values without shifting remaining rows
// Tags: Aspose.Cells delete blank rows | reverse row traversal Aspose.Cells | MaxDataRow based row removal .NET | Excel worksheet empty row cleanup | cell null detection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, determines the last used row and column via MaxDataRow/MaxDataColumn, iterates from the bottom row upward, checks each cell for a non‑null value, deletes rows that contain no data, and saves the cleaned workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row and column that contain data
            int lastRow = sheet.Cells.MaxDataRow;
            int lastCol = sheet.Cells.MaxDataColumn;

            // Loop from the last row upwards to avoid index shifting after deletions
            for (int row = lastRow; row >= 0; row--)
            {
                bool hasData = false;

                // Check each cell in the current row up to the last used column
                for (int col = 0; col <= lastCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // If any cell in the row has a non‑null value, the row is not empty
                    if (cell.Value != null)
                    {
                        hasData = true;
                        break; // No need to check further cells in this row
                    }
                }

                // Delete the row if it contains no data
                if (!hasData)
                {
                    sheet.Cells.DeleteRows(row, 1);
                }
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Processing complete. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
