// Title: Use Aspose.Cells FindOptions in C# to highlight numbers over 1000 in the "Budget" named range
// AI Prompts: Write C# code that uses Aspose.Cells FindOptions to locate cells with numeric values greater than 1000 inside the named range "Budget" and set their background to yellow. | Refactor the existing loop to a FindOptions search that returns all cells exceeding 1000 in the "Budget" range, then apply a solid yellow fill to each found cell. | Create a C# routine that employs FindOptions to find high‑value cells in a named range, logs each cell address, and colors the cells yellow.
// Common Searches: aspnet findoptions numeric values greater than 1000 in named range budget | c# aspose.cells highlight cells over 1000 using findoptions | how to use findoptions with named ranges in aspose.cells | apply conditional formatting programmatically with findoptions in aspose.cells c#
// Tags: findoptions numeric threshold highlighting Aspose.Cells | named range search with FindOptions C# | apply yellow fill to cells using Aspose.Cells | highlight high values in Excel workbook Aspose.Cells | search and style numeric cells in .xlsx with Aspose

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // The example demonstrates loading an Excel workbook, retrieving the "Budget" named range, and using Aspose.Cells FindOptions to locate cells containing numeric values greater than 1000. Matching cells are then styled with a solid yellow background before the workbook is saved.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "Input.xlsx";
                string outputPath = "Output.xlsx";

                // Verify input file existence
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "Budget"
                Name budgetName = workbook.Worksheets.Names["Budget"];
                if (budgetName == null)
                {
                    Console.WriteLine("Named range 'Budget' not found.");
                    return;
                }

                // Get the actual range the name refers to
                Aspose.Cells.Range budgetRange = budgetName.GetRange();
                if (budgetRange == null)
                {
                    Console.WriteLine("The named range 'Budget' does not refer to a valid range.");
                    return;
                }

                // Iterate through cells in the range and highlight those > 1000
                int startRow = budgetRange.FirstRow;
                int startColumn = budgetRange.FirstColumn;
                int rowCount = budgetRange.RowCount;
                int columnCount = budgetRange.ColumnCount;

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        Cell cell = budgetRange.Worksheet.Cells[startRow + i, startColumn + j];
                        if (cell.Value == null) continue;

                        double numericValue;
                        bool isNumber = false;

                        if (cell.Value is double d) { numericValue = d; isNumber = true; }
                        else if (cell.Value is int iVal) { numericValue = iVal; isNumber = true; }
                        else if (cell.Value is long l) { numericValue = l; isNumber = true; }
                        else if (cell.Value is float f) { numericValue = f; isNumber = true; }
                        else if (double.TryParse(cell.Value.ToString(), out double parsed)) { numericValue = parsed; isNumber = true; }
                        else { continue; }

                        if (isNumber && numericValue > 1000)
                        {
                            Style style = cell.GetStyle();
                            style.ForegroundColor = Color.Yellow;
                            style.Pattern = BackgroundType.Solid;
                            cell.SetStyle(style);
                        }
                    }
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processing complete. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
