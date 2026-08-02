// Title: Scan Excel workbook for OFFSET formulas and flag volatility using Aspose.Cells C#
// Description: A C# utility that loads an Excel file with Aspose.Cells, walks the used range of each sheet, identifies formulas containing the OFFSET function, and logs the worksheet, cell address, full formula, and its volatile nature.
// Keywords: Aspose.Cells | C# | OFFSET function | volatile Excel formula | formula scanner | Excel audit | detect Excel functions | performance optimization | Excel workbook analysis
// Common Searches: find OFFSET formulas with Aspose.Cells | list volatile functions in Excel using C# | scan workbook for specific formulas Aspose.Cells | detect OFFSET in Excel programmatically | C# code to identify volatile Excel formulas
// Developer Intent: Locate every formula that uses OFFSET in a workbook and indicate that it is a volatile function.
// Use Cases: Generate an audit report of all OFFSET formulas before publishing a workbook to improve calculation speed. | Integrate the scanner into a CI/CD pipeline to block builds that introduce new volatile OFFSET formulas. | Create a summary sheet or CSV file listing cells with OFFSET for documentation and refactoring.
// AI Prompts: Write a C# method that returns a collection of Cell objects whose formulas contain OFFSET using Aspose.Cells. | Extend the scanner to also detect INDIRECT and TODAY functions and classify each as volatile or non‑volatile. | Refactor the code to export the detection results to a CSV file instead of writing to the console.

using System;
using Aspose.Cells;

namespace AsposeCellsOffsetFormulaScanner
{
    // A C# utility that loads an Excel file with Aspose.Cells, walks the used range of each sheet, identifies formulas containing the OFFSET function, and logs the worksheet, cell address, full formula, and its volatile nature.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be scanned.
            // You can pass the path as a command‑line argument or modify the string directly.
            string workbookPath = args.Length > 0 ? args[0] : "InputWorkbook.xlsx";

            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to avoid scanning empty cells.
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell within the used range.
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only formula cells.
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Check for the OFFSET function (case‑insensitive).
                            if (formula != null && formula.IndexOf("OFFSET(", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // OFFSET is a volatile function – its result can change whenever any
                                // cell in the workbook recalculates, even if the referenced range does not change.
                                Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cell.Name}, Formula: {formula}, Volatile: Yes");
                            }
                        }
                    }
                }
            }

            // No need to save the workbook because this operation is read‑only.
        }
    }
}
