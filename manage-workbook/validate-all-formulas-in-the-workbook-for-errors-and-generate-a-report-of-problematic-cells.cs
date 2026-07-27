using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace FormulaValidationExample
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            string inputFile = "input.xlsx";
            Workbook workbook = new Workbook(inputFile);

            // Prepare a list to hold error information
            List<string> errorReport = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all used cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        try
                        {
                            // Attempt to calculate the formula directly.
                            // If the formula is invalid, an exception will be thrown.
                            object result = sheet.CalculateFormula(cell.Formula);
                            // Optionally, you could store successful results here.
                        }
                        catch (Exception ex)
                        {
                            // Record the cell address, worksheet name and the exception message
                            string cellAddress = cell.Name; // e.g., "A1"
                            string message = $"{sheet.Name}!{cellAddress}: {ex.Message}";
                            errorReport.Add(message);
                        }
                    }
                }
            }

            // Write the report to a text file
            string reportFile = "FormulaErrorsReport.txt";
            File.WriteAllLines(reportFile, errorReport);

            // Optionally, calculate all formulas (ignoring errors) after validation
            CalculationOptions calcOptions = new CalculationOptions { IgnoreError = true };
            workbook.CalculateFormula(calcOptions);

            // Save the workbook (if any changes were made)
            string outputFile = "output.xlsx";
            workbook.Save(outputFile);
        }
    }
}