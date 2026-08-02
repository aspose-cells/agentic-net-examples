// Title: Freeze TODAY() in Excel: Replace with a Fixed Date using Aspose.Cells for C#/.NET
// Description: Loads an Excel workbook with Aspose.Cells, scans every worksheet, detects cells that contain the TODAY() formula, substitutes each formula with a predefined static DateTime (e.g., 2023‑08‑01), and saves the workbook—effectively freezing dynamic date calculations.
// Keywords: Aspose.Cells | C# Excel automation | replace TODAY() | static date | freeze calculations | Excel formula replacement | Aspose.Cells .NET example | fixed date in workbook | remove dynamic date function | Excel automation C# | date function freeze | Aspose.Cells GitHub
// Common Searches: Aspose.Cells replace TODAY() with static date | C# freeze Excel TODAY function | How to convert TODAY() to constant date in .NET | Aspose.Cells fix dynamic date formulas | Replace dynamic date in Excel using C#
// Developer Intent: Replace all TODAY() formulas with a predetermined static date in an Excel workbook.
// Use Cases: Generate month‑end reports that retain the same date after distribution. | Archive compliance‑required workbooks where dates must remain unchanged. | Create Excel templates for regulatory filing with locked dates. | Capture data snapshots for audit trails without future recalculation. | Distribute Excel files to clients while preventing live date updates.
// AI Prompts: Write C# code with Aspose.Cells that finds and replaces TODAY() formulas with a given DateTime. | Show how to iterate through every cell in a workbook and substitute dynamic date functions with a constant value while preserving other formulas. | Explain steps to freeze dynamic date calculations in an Excel file using Aspose.Cells for .NET. | Provide a method to bulk replace TODAY() across multiple worksheets in a workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook with Aspose.Cells, scans every worksheet, detects cells that contain the TODAY() formula, substitutes each formula with a predefined static DateTime (e.g., 2023‑08‑01), and saves the workbook—effectively freezing dynamic date calculations.
    public class FreezeTodayFunctionDemo
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Static date to replace TODAY()
                DateTime staticDate = new DateTime(2023, 8, 1);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Check if the cell contains a formula that uses TODAY()
                            if (cell.IsFormula &&
                                !string.IsNullOrEmpty(cell.Formula) &&
                                cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Replace the formula with the static date value
                                cell.PutValue(staticDate);
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Program entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
