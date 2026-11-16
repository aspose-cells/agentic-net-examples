// Title: C# example: delete rows above a target row and set FreezePanes with Aspose.Cells
// Common Searches: C# Aspose.Cells delete rows above a specific row before freezing panes | How to remove leading rows in an Excel file and then apply FreezePanes using Aspose.Cells | Aspose.Cells code to delete top rows and freeze the first visible row in .NET | Programmatically adjust FreezePanes after deleting rows with Aspose.Cells | Example for deleting rows prior to a freeze point and setting FreezePanes in C#
// Tags: delete rows before freeze panes Aspose.Cells | set freeze panes after row deletion Aspose.Cells | Aspose.Cells remove top rows Excel | freeze first visible row Aspose.Cells C# | adjust freeze pane index Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, deletes all rows above the specified freeze row, updates the FreezePanes settings to freeze the new first row, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Example parameters – replace with actual values or pass via args
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            int freezeRow = 5; // 1‑based row that should become the frozen row after deletion

            try
            {
                ProcessWorkbook(inputPath, outputPath, freezeRow);
                Console.WriteLine("Workbook processed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        static void ProcessWorkbook(string inputPath, string outputPath, int freezeRow)
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook
                var workbook = new Workbook(inputPath);
                var sheet = workbook.Worksheets[0];

                // Delete rows above the intended freeze row
                int rowsToDelete = freezeRow - 1;
                if (rowsToDelete > 0)
                {
                    sheet.Cells.DeleteRows(0, rowsToDelete);
                }

                // Freeze the first visible row (now the intended row)
                // FreezePanes(row, column, frozenRows, frozenColumns)
                sheet.FreezePanes(1, 0, 1, 0); // freeze first row, no columns

                // Ensure output directory exists
                string outDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Rethrow to allow outer handler to report the error
                throw new ApplicationException($"Failed to process workbook: {ex.Message}", ex);
            }
        }
    }
}
