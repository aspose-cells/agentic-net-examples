using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class FormulaValidator
{
    static void Main()
    {
        // Load the workbook (replace with actual path)
        Workbook workbook = new Workbook("input.xlsx");

        // List to hold error information
        List<string> errorReport = new List<string>();

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Scan every cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        try
                        {
                            // Attempt to calculate the formula using the worksheet method
                            // This will throw if the formula is invalid or unsupported
                            object result = sheet.CalculateFormula(cell.Formula);
                            // Optionally you could store the result back, but it's not required for validation
                        }
                        catch (Exception ex)
                        {
                            // Record the worksheet name, cell address, and error message
                            string cellAddress = cell.Name; // e.g., "A1"
                            errorReport.Add($"{sheet.Name}!{cellAddress}: {ex.Message}");
                        }
                    }
                }
            }
        }

        // Write the error report to a text file
        string reportPath = "FormulaErrorsReport.txt";
        File.WriteAllLines(reportPath, errorReport);

        Console.WriteLine($"Formula validation completed. Report saved to '{reportPath}'.");
    }
}