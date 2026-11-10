// Title: How to replace RAND() formulas with a fixed value in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells replace RAND() function with constant value in C# | Make Excel random formulas deterministic with Aspose.Cells .NET | Iterate through cells and change RAND() to 0.5 using Aspose.Cells | Recalculate workbook after modifying formulas in Aspose.Cells C# | Load workbook, replace RAND() formulas and save with Aspose.Cells
// Tags: RAND() formula replacement Aspose.Cells | static seed for Excel random function .NET | deterministic random calculation C# | cell iteration for formula modification Aspose | formula recalculation after edit Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file with Aspose.Cells, walks through every worksheet and cell, substitutes any RAND() occurrence in formulas with a fixed value (e.g., 0.5), forces a full formula recalculation, and saves the workbook to produce reproducible results.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells in the worksheet
                    foreach (Cell cell in cells)
                    {
                        // Process only formula cells
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                        {
                            // Replace RAND() with a static value (e.g., 0.5)
                            if (cell.Formula.Contains("RAND()"))
                            {
                                cell.Formula = cell.Formula.Replace("RAND()", "0.5");
                            }
                        }
                    }
                }

                // Recalculate formulas after modifications
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
