// Title: Expand an Excel ListObject by adding rows and refresh its named range with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to resize a ListObject named 'Table1' by a specified number of rows and automatically set the RefersTo property of the named range 'MyRange' to the new table address. | After increasing the size of an Excel table, programmatically rebuild the A1‑style address and assign it to an existing workbook named range using C# and Aspose.Cells.
// Common Searches: C# Aspose.Cells how to add rows to an existing Excel table and keep a named range updated | Resize ListObject and update named range RefersTo with Aspose.Cells .NET | Expand Excel table programmatically and synchronize named range using Aspose.Cells C# example
// Tags: resize ListObject Aspose.Cells | update named range RefersTo C# | programmatically extend Excel ListObject | synchronize named range with expanded table | Aspose.Cells table expansion example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The sample loads a workbook, locates a ListObject called 'Table1', expands it by a defined number of rows using the Resize method, rebuilds the A1‑style address of the resized table, and updates the RefersTo property of the named range 'MyRange' so it points to the new range before saving the file.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the table (ListObject) by its name
                ListObject table = worksheet.ListObjects["Table1"]; // replace with your table name
                if (table == null)
                {
                    Console.WriteLine("Table 'Table1' not found.");
                    return;
                }

                // Number of new rows to add to the table
                int rowsToAdd = 5;

                // Calculate new bottom row index for the table after expansion
                int startRow = table.StartRow;
                int startColumn = table.StartColumn;
                int endRow = table.EndRow;
                int endColumn = table.EndColumn;

                int newEndRow = endRow + rowsToAdd;
                int totalRows = newEndRow - startRow + 1;
                int totalColumns = endColumn - startColumn + 1;

                // Resize the table to include the new rows
                try
                {
                    table.Resize(startRow, startColumn, totalRows, totalColumns, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to resize table: {ex.Message}");
                    return;
                }

                // Update the named range that references the table (if it exists)
                try
                {
                    Name namedRange = workbook.Worksheets.Names["MyRange"]; // replace with your named range
                    if (namedRange != null)
                    {
                        // Build A1 style address for the table range
                        string startAddr = CellsHelper.CellIndexToName(table.StartRow, table.StartColumn);
                        string endAddr = CellsHelper.CellIndexToName(table.EndRow, table.EndColumn);
                        string tableAddress = $"{startAddr}:{endAddr}";

                        // Set the RefersTo property (include leading '=' and sheet name)
                        namedRange.RefersTo = $"={worksheet.Name}!{tableAddress}";
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to update named range: {ex.Message}");
                }

                // Ensure output directory exists
                string outDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
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
}
