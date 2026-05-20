using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class FormulaValidator
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Collection to store information about cells with formula errors
        List<string> errorReport = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Scan every cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        try
                        {
                            // Evaluate the formula; any exception indicates a problem
                            sheet.CalculateFormula(cell.Formula);
                        }
                        catch (Exception ex)
                        {
                            // Record the worksheet name, cell address, and error message
                            errorReport.Add($"{sheet.Name}!{cell.Name}: {ex.Message}");
                        }
                    }
                }
            }
        }

        // Output the report to a text file
        File.WriteAllLines("FormulaErrorsReport.txt", errorReport);

        // Optional console feedback
        Console.WriteLine($"Formula validation completed. Problems found: {errorReport.Count}");
    }
}