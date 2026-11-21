// Title: How to remove external workbook references from Excel formulas while preserving calculated values using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace external links in formulas with static values | remove external workbook references from Excel file before distribution using Aspose.Cells | clear formulas that contain external references in .xlsx with Aspose.Cells .NET | preserve calculated results while deleting external links in Excel using Aspose.Cells | iterate cells to detect '[' in formula and remove external reference Aspose.Cells C#
// Tags: Aspose.Cells external reference removal | clear external formulas while preserving values | Excel workbook formula cleanup with Aspose.Cells | C# iterate worksheet cells to detect external links | save workbook without external links Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, recalculates all formulas, then walks through each worksheet and cell to find formulas containing '[' (the marker for external references). For each such cell it captures the current calculated value, replaces the formula with that value, clears the formula string, and finally saves the cleaned workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas to ensure values are up‑to‑date
            workbook.CalculateFormula();

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only cells that contain a formula
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // External references are indicated by '[' in the formula string
                            if (!string.IsNullOrEmpty(formula) && formula.Contains("["))
                            {
                                // Preserve the calculated result, then remove the formula
                                object calculatedValue = cell.Value;
                                cell.PutValue(calculatedValue);
                                cell.Formula = string.Empty; // clear the external formula
                            }
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
