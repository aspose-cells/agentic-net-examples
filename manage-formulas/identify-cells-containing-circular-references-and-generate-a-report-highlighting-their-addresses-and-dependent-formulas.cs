// Title: Identify circular reference cells in an Excel workbook and generate a detailed report with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that scans every worksheet, attempts to calculate each formula cell, catches circular reference exceptions, and records the sheet name, cell address, and formula into a new Excel file. | Create a C# routine with Aspose.Cells that builds a list of cells causing circular references, includes the exception message, and exports the list as a CSV file. | Develop a C# program that processes an input workbook, skips non‑formula cells, logs unexpected calculation errors, and produces a circular reference summary workbook using Aspose.Cells.
// Common Searches: how to find circular reference cells in Excel using Aspose.Cells C# | Aspose.Cells generate report of formula errors in .xlsx file | C# detect and list circular references in workbook with Aspose.Cells calculation options | export circular reference details to new Excel file using Aspose.Cells library | handle circular reference exception while calculating formulas in Aspose.Cells
// Tags: detect circular references Aspose.Cells | export circular reference report Excel C# | calculate formulas with exception handling Aspose.Cells | list formula cells causing circular reference Aspose.Cells | generate error summary workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace CircularReferenceReport
{
    // The program loads an input workbook, iterates through every worksheet and formula cell, attempts to calculate each cell, catches exceptions containing "circular reference", records the sheet name, cell address and formula, then writes this information into a new workbook saved as CircularReferenceReport.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "CircularReferenceReport.xlsx";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that needs to be inspected.
                Workbook sourceWorkbook = new Workbook(inputPath);

                // List to hold information about cells that cause circular references.
                List<(string SheetName, string CellAddress, string Formula)> circularCells =
                    new List<(string, string, string)>();

                // Options required for cell calculation.
                CalculationOptions calcOptions = new CalculationOptions();

                // Iterate through all worksheets and cells.
                foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Loop through each cell that contains a formula.
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula)
                        {
                            try
                            {
                                // Attempt to calculate the formula for this cell.
                                // If a circular reference exists, Aspose.Cells throws an exception.
                                cell.Calculate(calcOptions);
                            }
                            catch (Exception ex)
                            {
                                // Detect circular reference by checking the exception message.
                                if (!string.IsNullOrEmpty(ex.Message) &&
                                    ex.Message.IndexOf("circular reference", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // Record the sheet name, cell address and its formula.
                                    circularCells.Add((sheet.Name, cell.Name, cell.Formula));
                                }
                                else
                                {
                                    // Log unexpected errors for this cell but continue processing.
                                    Console.WriteLine($"Warning: Unable to calculate cell {cell.Name} in sheet \"{sheet.Name}\": {ex.Message}");
                                }
                            }
                        }
                    }
                }

                // Create a new workbook to hold the report.
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];
                Cells reportCells = reportSheet.Cells;

                // Write header row.
                reportCells[0, 0].PutValue("Worksheet");
                reportCells[0, 1].PutValue("Cell Address");
                reportCells[0, 2].PutValue("Formula");

                // Populate the report with the collected circular reference data.
                int reportRow = 1;
                foreach (var entry in circularCells)
                {
                    reportCells[reportRow, 0].PutValue(entry.SheetName);
                    reportCells[reportRow, 1].PutValue(entry.CellAddress);
                    reportCells[reportRow, 2].PutValue(entry.Formula);
                    reportRow++;
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the report workbook.
                reportWorkbook.Save(outputPath);
                Console.WriteLine($"Circular reference report saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
