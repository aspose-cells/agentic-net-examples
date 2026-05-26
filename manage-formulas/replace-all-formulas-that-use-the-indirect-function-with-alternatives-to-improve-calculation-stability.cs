using System;
using Aspose.Cells;

namespace AsposeCellsIndirectReplacement
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "InputWorkbook.xlsx";
            Workbook wb = new Workbook(inputPath);

            // Ensure all formulas are calculated so we can retrieve their results
            wb.CalculateFormula();

            // Iterate through each worksheet
            foreach (Worksheet sheet in wb.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Scan every used cell in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a formula
                    if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                    {
                        // Check if the formula uses the INDIRECT function
                        if (cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Retrieve the already‑calculated value of the cell
                            object calculatedValue = cell.Value;

                            // Replace the volatile INDIRECT formula with the static value
                            // First clear the formula, then put the calculated value back
                            cell.Formula = string.Empty;
                            cell.PutValue(calculatedValue);
                        }
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            string outputPath = "OutputWorkbook.xlsx";
            wb.Save(outputPath);
        }
    }
}