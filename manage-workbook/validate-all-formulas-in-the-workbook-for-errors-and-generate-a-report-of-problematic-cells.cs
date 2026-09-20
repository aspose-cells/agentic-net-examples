// Title: Detect and report Excel formula errors across all worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, forces full formula calculation, scans every worksheet for cells where the result type is error, and outputs a plain‑text report listing sheet name, cell address, and the error value. | Adapt the formula‑validation routine to generate a CSV file instead of plain text, including the specific Excel error codes (e.g., #DIV/0!, #VALUE!) for each problematic cell. | Create a reusable C# method `ValidateFormulas(string workbookPath, string reportPath, bool csv = false)` that returns a list of objects containing worksheet, address, and error, and writes the chosen report format.
// Common Searches: how to programmatically find #DIV/0 errors in an Excel workbook using Aspose.Cells C# | Aspose.Cells C# example to generate a list of cells with formula errors | C# iterate through all worksheets and detect formula calculation errors with Aspose.Cells | export formula error details from an .xlsx file to a text or CSV report using Aspose.Cells
// Tags: Aspose.Cells formula error detection | C# iterate worksheets used range | export formula error report to text | calculate workbook formulas Aspose.Cells | list error cells Excel .xlsx C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace FormulaValidation
{
    // Loads an Excel workbook with Aspose.Cells, forces formula calculation, walks through each worksheet's used range, captures cells whose formula result is an error, and writes a plain‑text report containing worksheet name, cell address, and error value.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output paths (adjust as needed)
            string inputFile = "input.xlsx";
            string reportFile = "FormulaValidationReport.txt";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Ensure all formulas are calculated so errors can be detected
                workbook.CalculateFormula();

                // List to hold information about cells with formula errors
                List<string> errorCells = new List<string>();

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                    // Determine the bounds of the used range
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startColumn = usedRange.FirstColumn;
                    int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Iterate through each cell in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startColumn; col <= endColumn; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Process only cells that contain a formula
                            if (cell.IsFormula)
                            {
                                // If the formula result is an error, record details
                                if (cell.Type == CellValueType.IsError)
                                {
                                    string cellAddress = cell.Name; // e.g., "A1"
                                    string sheetName = sheet.Name;
                                    string errorInfo = $"{sheetName}!{cellAddress} - Error: {cell.Value}";
                                    errorCells.Add(errorInfo);
                                }
                            }
                        }
                    }
                }

                // Generate the report
                try
                {
                    using (StreamWriter writer = new StreamWriter(reportFile))
                    {
                        writer.WriteLine("Formula Validation Report");
                        writer.WriteLine($"Generated on: {DateTime.Now}");
                        writer.WriteLine();

                        if (errorCells.Count == 0)
                        {
                            writer.WriteLine("No formula errors were detected.");
                        }
                        else
                        {
                            writer.WriteLine("Problematic cells with formula errors:");
                            foreach (string line in errorCells)
                            {
                                writer.WriteLine(line);
                            }
                        }
                    }

                    Console.WriteLine($"Validation complete. Report saved to '{reportFile}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write report file: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}
