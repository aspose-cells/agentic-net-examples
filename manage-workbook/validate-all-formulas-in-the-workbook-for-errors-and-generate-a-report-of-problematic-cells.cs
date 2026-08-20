// Title: C# – Validate Excel Formulas and Create an Error Report with Aspose.Cells
// Description: Load a workbook, calculate all formulas while ignoring errors, scan every used cell for error values, record the worksheet name, cell address, and error text on a new "FormulaErrors" sheet, and save the updated file. This example shows how to use Aspose.Cells CalculationOptions, CellValueType.IsError, and worksheet iteration to produce a comprehensive formula‑error summary.
// Keywords: Aspose.Cells formula validation | C# Excel error report | detect #DIV/0! Aspose | calculate formulas ignore errors | list error cells Aspose.Cells | Excel workbook diagnostics .NET | generate formula error sheet | Aspose.Cells CalculationOptions
// Common Searches: Aspose.Cells how to find cells with errors | C# generate Excel formula error summary | calculate formulas without throwing exceptions Aspose | iterate through workbook cells to detect #N/A | create error report worksheet using Aspose.Cells
// Developer Intent: Find every cell that returns an error after formula calculation and compile a consolidated report worksheet.
// Use Cases: Create a diagnostic tab that lists all #DIV/0!, #N/A, #VALUE! and similar errors before sharing the workbook. | Automate pre‑release validation of financial models to ensure no error values remain. | Integrate formula‑error detection into CI/CD pipelines that generate or modify Excel files.
// AI Prompts: Generate C# code with Aspose.Cells that scans a workbook for cells of type IsError and writes the sheet name, address, and error message to a new sheet called "FormulaErrors". | Explain how CalculationOptions.IgnoreError allows safe formula evaluation while collecting error information in Aspose.Cells. | Provide step‑by‑step guidance for building a summary worksheet of formula errors and saving the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaValidation
{
    // Load a workbook, calculate all formulas while ignoring errors, scan every used cell for error values, record the worksheet name, cell address, and error text on a new "FormulaErrors" sheet, and save the updated file. This example shows how to use Aspose.Cells CalculationOptions, CellValueType.IsError, and worksheet iteration to produce a comprehensive formula‑error summary.
    class Program
    {
        static void Main()
        {
            // ---------- Create / Load ----------
            // Load an existing workbook (replace with actual path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // ---------- Calculate all formulas ----------
            // Use calculation options that ignore errors so the calculation does not throw
            CalculationOptions calcOptions = new CalculationOptions { IgnoreError = true };
            workbook.CalculateFormula(calcOptions);

            // ---------- Validate formulas ----------
            // Create a new worksheet to hold the error report
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "FormulaErrors";

            // Write header row
            Cells reportCells = reportSheet.Cells;
            reportCells[0, 0].PutValue("Worksheet");
            reportCells[0, 1].PutValue("Cell");
            reportCells[0, 2].PutValue("Error");

            int reportRow = 1; // start after header

            // Iterate through all worksheets and their used cells
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the report sheet itself
                if (ws.Name == reportSheet.Name) continue;

                Cells cells = ws.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Check if the cell contains an error value
                        if (cell != null && cell.Type == CellValueType.IsError)
                        {
                            // Record the problematic cell
                            reportCells[reportRow, 0].PutValue(ws.Name);
                            reportCells[reportRow, 1].PutValue(cell.Name);
                            reportCells[reportRow, 2].PutValue(cell.StringValue); // e.g., "#DIV/0!"
                            reportRow++;
                        }
                    }
                }
            }

            // ---------- Save ----------
            // Save the workbook with the error report (replace with desired output path)
            string outputPath = "output_with_errors_report.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Formula validation completed. Report saved to: " + outputPath);
        }
    }
}
