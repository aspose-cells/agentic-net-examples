// Title: Duplicate an Excel ListObject to a new worksheet while preserving styles, formulas, and totals using Aspose.Cells for .NET
// Common Searches: how to copy a ListObject to another sheet with formatting using Aspose.Cells C# | Aspose.Cells duplicate Excel table preserving formulas and totals row | C# copy Excel table to new worksheet keeping cell styles with Aspose.Cells | clone a worksheet table and retain total row setting in Aspose.Cells .NET | copy range of a table and recreate ListObject on another sheet Aspose.Cells
// Tags: copy ListObject with styles Aspose.Cells | duplicate Excel table preserving formulas C# | clone worksheet table including totals row Aspose.Cells | create new ListObject from copied range Aspose.Cells | transfer table data between worksheets Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableCopy
{
    // The example loads a workbook, selects the first ListObject on the source sheet, copies each cell's value, style, and formula to a new worksheet, creates a matching ListObject on the destination sheet, copies the totals‑row setting, and saves the workbook with the duplicated table.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Source worksheet that contains the table to duplicate
                Worksheet srcSheet = workbook.Worksheets[0]; // or use name: workbook.Worksheets["Source"]

                // Assume the first table (ListObject) on the source sheet is the one to copy
                if (srcSheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables found on the source worksheet.");
                    return;
                }

                ListObject srcTable = srcSheet.ListObjects[0];

                // Determine the source table bounds
                int srcStartRow = srcTable.StartRow;
                int srcStartColumn = srcTable.StartColumn;
                int srcEndRow = srcTable.EndRow;
                int srcEndColumn = srcTable.EndColumn;

                int rowCount = srcEndRow - srcStartRow + 1;
                int columnCount = srcEndColumn - srcStartColumn + 1;

                // Create (or get) the destination worksheet where the duplicate will be placed
                Worksheet destSheet = workbook.Worksheets.Add("TableCopy");

                // Destination top‑left cell (e.g., A1)
                int destStartRow = 0;
                int destStartColumn = 0;

                // Copy the cells (values, formulas, formats) from the source range to the destination range
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < columnCount; c++)
                    {
                        Cell srcCell = srcSheet.Cells[srcStartRow + r, srcStartColumn + c];
                        Cell destCell = destSheet.Cells[destStartRow + r, destStartColumn + c];
                        destCell.PutValue(srcCell.Value);
                        destCell.SetStyle(srcCell.GetStyle());
                        destCell.Formula = srcCell.Formula;
                    }
                }

                // Add a new ListObject (table) on the destination sheet with the same size as the source table
                int destTableIndex = destSheet.ListObjects.Add(
                    destStartRow,          // first row of the table
                    destStartColumn,       // first column of the table
                    rowCount,              // number of rows (including header/total)
                    columnCount,           // number of columns
                    true);                 // has header row

                ListObject destTable = destSheet.ListObjects[destTableIndex];

                // Preserve the total row setting
                destTable.ShowTotals = srcTable.ShowTotals;

                // Save the workbook with the duplicated table
                try
                {
                    // Ensure the output directory exists
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputPath);
                    Console.WriteLine($"Table duplicated successfully. Output saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save the workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
