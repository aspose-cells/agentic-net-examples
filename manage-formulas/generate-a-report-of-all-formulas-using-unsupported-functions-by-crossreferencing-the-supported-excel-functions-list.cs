// Title: C# – Generate a Report of Unsupported Excel Formulas Using Aspose.Cells (.NET)
// Description: Loads an Excel workbook with Aspose.Cells, iterates through every worksheet, cell and named range, detects formulas that contain unsupported (custom) functions via the HasCustomFunction property, and writes a detailed text report with worksheet name, cell address and formula. Includes robust file‑existence checks and graceful error handling.
// Keywords: Aspose.Cells unsupported functions | detect custom Excel formulas .NET | HasCustomFunction property | C# scan workbook for invalid formulas | generate Excel formula audit report | Aspose.Cells formula validation | Excel unsupported functions USA | Excel formula compliance Europe
// Common Searches: how to list unsupported Excel functions with Aspose.Cells | C# code to find custom formulas in a workbook | Aspose.Cells generate report of invalid formulas | detect unsupported functions in named ranges Aspose.Cells | audit Excel file for formulas not supported by Aspose
// Developer Intent: Identify every formula that uses a function not supported by Aspose.Cells and produce a readable report for further analysis or remediation.
// Use Cases: Pre‑processing audit to ensure a workbook can be converted to PDF without formula errors. | Compliance reporting for large spreadsheets that must only contain supported functions. | Automated validation of named ranges before performing bulk data extraction or migration.
// AI Prompts: Create C# code that exports the unsupported‑function report to CSV instead of TXT, including worksheet index and cell address. | Enhance the program to count each unique unsupported function and add a summary section at the end of the report. | Add logging for worksheets that fail to load and guarantee the temporary cell used for named‑range checks is always cleared, even on exceptions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook with Aspose.Cells, iterates through every worksheet, cell and named range, detects formulas that contain unsupported (custom) functions via the HasCustomFunction property, and writes a detailed text report with worksheet name, cell address and formula. Includes robust file‑existence checks and graceful error handling.
    public class UnsupportedFormulasReport
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Report file path
            string reportPath = "UnsupportedFormulasReport.txt";

            try
            {
                using (StreamWriter writer = new StreamWriter(reportPath))
                {
                    writer.WriteLine("Unsupported Formulas Report");
                    writer.WriteLine($"Generated on: {DateTime.Now}");
                    writer.WriteLine(new string('=', 50));
                    writer.WriteLine();

                    // Scan each worksheet for cells containing unsupported/custom functions
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        foreach (Cell cell in sheet.Cells)
                        {
                            if (string.IsNullOrEmpty(cell.Formula))
                                continue;

                            if (cell.HasCustomFunction)
                            {
                                writer.WriteLine($"Worksheet: {sheet.Name}");
                                writer.WriteLine($"Cell     : {cell.Name}");
                                writer.WriteLine($"Formula  : {cell.Formula}");
                                writer.WriteLine(new string('-', 40));
                            }
                        }
                    }

                    // Check defined names (named ranges) for custom functions
                    foreach (Aspose.Cells.Name definedName in workbook.Worksheets.Names)
                    {
                        if (!string.IsNullOrEmpty(definedName.RefersTo) && definedName.RefersTo.StartsWith("="))
                        {
                            // Use a temporary cell to evaluate the formula
                            Cell tempCell = workbook.Worksheets[0].Cells["ZZ1"];
                            tempCell.Formula = definedName.RefersTo;

                            if (tempCell.HasCustomFunction)
                            {
                                // Output the defined name's reference (name property may not be available in some versions)
                                writer.WriteLine($"Defined Name Refers To: {definedName.RefersTo}");
                                writer.WriteLine(new string('-', 40));
                            }

                            // Clear the temporary cell to avoid side effects
                            tempCell.Formula = string.Empty;
                        }
                    }

                    writer.WriteLine();
                    writer.WriteLine("Report generation completed.");
                }

                Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to generate report: {ex.Message}");
            }
        }
    }
}
