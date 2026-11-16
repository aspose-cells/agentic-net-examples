// Title: Delete left‑side columns before applying FreezePanes with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells delete columns before freezing panes in Excel | How to remove left columns and set FreezePanes using Aspose.Cells .NET | Aspose.Cells example for deleting columns and applying FreezePanes programmatically | FreezePanes after column deletion with Aspose.Cells in a C# console app | Command‑line tool to trim columns and freeze rows with Aspose.Cells
// Tags: Aspose.Cells column deletion API | FreezePanes configuration Aspose.Cells C# | C# remove leftmost columns Excel worksheet | Aspose.Cells command‑line workbook manipulation | Excel freeze panes after column removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# console example loads an Excel workbook with Aspose.Cells, optionally deletes a specified number of columns from the left side of the first worksheet, applies FreezePanes at the given row (resetting the column index to zero), and saves the result. It demonstrates command‑line argument handling, column removal via DeleteColumns, and pane freezing with zero‑based indices.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Expected arguments: inputFile outputFile freezeRow freezeColumn
                if (args.Length < 4)
                {
                    Console.WriteLine("Usage: <inputFile> <outputFile> <freezeRow> <freezeColumn>");
                    return;
                }

                string inputFile = args[0];
                string outputFile = args[1];

                if (!int.TryParse(args[2], out int freezeRow))
                {
                    Console.WriteLine("Invalid freezeRow value.");
                    return;
                }

                if (!int.TryParse(args[3], out int freezeColumn))
                {
                    Console.WriteLine("Invalid freezeColumn value.");
                    return;
                }

                // Ensure the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputFile);
                var worksheet = workbook.Worksheets[0];

                // Delete columns to the left of the freeze column (if any)
                if (freezeColumn > 0)
                {
                    try
                    {
                        // DeleteColumns(startColumn, totalColumns, deleteData)
                        worksheet.Cells.DeleteColumns(0, freezeColumn, true);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error deleting columns: {ex.Message}");
                        return;
                    }
                }

                // Apply freeze panes (row and column indices are zero‑based)
                try
                {
                    // FreezePanes(row, column, totalRows, totalColumns)
                    worksheet.FreezePanes(freezeRow, 0, freezeRow, 0);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error applying freeze panes: {ex.Message}");
                    return;
                }

                // Ensure output directory exists
                try
                {
                    string outDir = Path.GetDirectoryName(outputFile);
                    if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                    {
                        Directory.CreateDirectory(outDir);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error preparing output directory: {ex.Message}");
                    return;
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
