// Title: C# Aspose.Cells example: detect formulas with undefined names and replace them with #REF! error
// Common Searches: C# Aspose.Cells replace #NAME? error with #REF! in an Excel workbook | How to detect formulas that reference undefined names using Aspose.Cells for .NET | Iterate all cells in a workbook and change #NAME? to #REF! with Aspose.Cells | Aspose.Cells find and fix undefined name errors in Excel files C# | Replace #NAME? with #REF! after calculating formulas using Aspose.Cells
// Tags: Aspose.Cells undefined name detection | C# Excel formula error replacement | Aspose.Cells error value conversion | Iterate workbook cells Aspose.Cells | Calculate formulas Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook with Aspose.Cells, forces formula calculation to expose errors, scans each worksheet's used range, identifies cells whose formulas evaluate to the "#NAME?" error, replaces those formulas with a "#REF!" error value, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas so that error values are materialized
            workbook.CalculateFormula();

            // Iterate through each worksheet
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
                            // After calculation, an undefined name results in a #NAME? error string
                            string valueStr = cell.Value?.ToString();
                            if (!string.IsNullOrEmpty(valueStr) &&
                                valueStr.Equals("#NAME?", StringComparison.OrdinalIgnoreCase))
                            {
                                // Replace the formula with a #REF! error representation
                                cell.PutValue("#REF!");
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
