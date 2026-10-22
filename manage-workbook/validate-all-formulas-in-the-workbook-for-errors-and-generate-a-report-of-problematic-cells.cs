// Title: Validate all formulas in an Excel workbook and generate a text report of cells with errors using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to find cells that return #DIV/0! or other formula errors | C# generate a text report of formula errors from an Excel file with Aspose.Cells | Validate workbook formulas and export error list to a .txt file using Aspose.Cells | Iterate through worksheets and cells to detect formula error values in Aspose.Cells C# | Force calculation of all formulas before checking for errors with Aspose.Cells
// Tags: Aspose.Cells calculate formulas | Aspose.Cells detect formula errors | Aspose.Cells export error list to txt | Aspose.Cells iterate worksheets cells | Aspose.Cells cell error string handling

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, forces calculation of all formulas, scans each worksheet for cells whose formula results are error strings (e.g., "#DIV/0!"), records the sheet name, cell address and error type, and writes the collected information to a plain‑text report.
class Program
{
    static void Main()
    {
        // Define input and output paths
        string inputPath = "input.xlsx";
        string reportPath = "FormulaErrorsReport.txt";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Force calculation of all formulas so that error values are materialized
            workbook.CalculateFormula();

            // List to hold information about cells with formula errors
            List<string> errorReport = new List<string>();

            // Iterate through each worksheet in the workbook
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
                            // In many Aspose.Cells versions, error values are represented as strings like "#DIV/0!"
                            if (cell.Value is string strValue && strValue.StartsWith("#"))
                            {
                                string cellAddress = cell.Name;   // e.g., "B5"
                                string sheetName = sheet.Name;    // worksheet name
                                string errorType = strValue;      // error string
                                errorReport.Add($"{sheetName}!{cellAddress}: {errorType}");
                            }
                        }
                    }
                }
            }

            // Write the report to a text file (ensure directory exists)
            try
            {
                File.WriteAllLines(reportPath, errorReport);
                Console.WriteLine($"Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write report file: {ex.Message}");
            }

            Console.WriteLine($"Validation complete. {errorReport.Count} problematic cells found.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
