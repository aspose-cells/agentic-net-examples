// Title: C# Example: Delete Empty Columns from an XLS Workbook and Export Cleaned Data to JSON using Aspose.Cells
// Common Searches: how to programmatically remove blank columns from an XLS file in C# with Aspose.Cells | Aspose.Cells C# code to clean Excel worksheets and save as JSON | delete empty columns in each worksheet before converting Excel to JSON in .NET | remove columns with no data using Aspose.Cells and export workbook to JSON | C# Aspose.Cells example for XLS to JSON conversion after column cleanup
// Tags: remove empty columns Aspose.Cells C# | XLS workbook column cleanup Aspose.Cells | Aspose.Cells export workbook to JSON | delete blank columns Excel worksheet .NET | Aspose.Cells save as JSON after cleanup

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the input XLS file exists, loads it with Aspose.Cells, iterates through each worksheet, scans columns from right to left, deletes any column where all cells are null or have no value, and then saves the cleaned workbook as a JSON file.
class Program
{
    static void Main()
    {
        // Paths for input workbook and output JSON
        string inputPath = "input.xls";
        string outputPath = "output.json";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the current maximum column index
                int maxColumn = sheet.Cells.MaxColumn;

                // Scan columns from right to left to safely delete empty columns
                for (int col = maxColumn; col >= 0; col--)
                {
                    bool columnIsEmpty = true;

                    // Check each cell in the current column
                    for (int row = 0; row <= sheet.Cells.MaxRow; row++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Determine if the cell is empty (no value and type is Null)
                        if (cell.Value != null || cell.Type != CellValueType.IsNull)
                        {
                            columnIsEmpty = false;
                            break; // No need to continue scanning this column
                        }
                    }

                    // Delete the column if it contains no data
                    if (columnIsEmpty)
                    {
                        sheet.Cells.DeleteColumn(col);
                    }
                }
            }

            // Export the cleaned workbook to JSON format
            workbook.Save(outputPath, SaveFormat.Json);
            Console.WriteLine($"Workbook successfully saved as JSON to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
