// Title: Remove a column from an Excel ListObject and clear its formulas using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that clears all formulas in a specified ListObject column and then deletes that column from the table. | Show how to iterate over a ListObject's data range in Aspose.Cells to purge formulas before removing the column at index 1. | Provide a .NET example that safely deletes a table column in an Excel worksheet while ensuring no formula remnants remain in other rows.
// Common Searches: Aspose.Cells how to delete a column from a ListObject without leaving formulas | C# remove Excel table column and purge formulas using Aspose.Cells | remove formulas in a ListObject column before deletion Aspose.Cells .NET | prevent formula propagation after deleting a table column with Aspose.Cells | remove ListObject column programmatically in C# and clean up formulas
// Tags: remove ListObject column Aspose.Cells C# | purge column formulas Aspose.Cells | delete Excel table column without residual formulas .NET | Aspose.Cells ListObject column removal data integrity | sanitize formulas before ListObject column deletion C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, verifies a ListObject exists, clears any formulas in a chosen data column of the table, removes that column from the ListObject, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Access the first ListObject
            ListObject table = sheet.ListObjects[0];

            // Zero‑based index of the column to remove (e.g., 1 = second column)
            int columnIndex = 1;

            // Clear any formulas in the target column before removal
            AsposeRange dataRange = table.DataRange; // Data rows only (excludes header)
            int startRow = dataRange.FirstRow;
            int endRow = dataRange.FirstRow + dataRange.RowCount - 1;
            int startColumn = dataRange.FirstColumn;

            for (int row = startRow; row <= endRow; row++)
            {
                int col = startColumn + columnIndex;
                Cell cell = sheet.Cells[row, col];
                if (cell.IsFormula)
                {
                    // Remove formula/value
                    cell.PutValue(string.Empty);
                }
            }

            // Remove the column from the ListObject
            if (columnIndex >= 0 && columnIndex < table.ListColumns.Count)
            {
                table.ListColumns.RemoveAt(columnIndex);
            }
            else
            {
                Console.WriteLine("Column index is out of range.");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
