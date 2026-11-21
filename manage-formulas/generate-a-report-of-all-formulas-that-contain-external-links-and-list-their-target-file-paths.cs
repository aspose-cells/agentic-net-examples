// Title: Generate a report of Excel formulas containing external links and their target file paths using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells list cells with external workbook references | how to extract external file paths from Excel formulas using Aspose.Cells | create a report of external link formulas in an .xlsx file with Aspose.Cells | iterate through worksheets and cells to find external links Aspose.Cells C# | save external link details to a new Excel workbook using Aspose.Cells .NET
// Tags: extract external link formulas Aspose.Cells | enumerate worksheet cells with external references C# | generate formula link report Excel Aspose.Cells | export external link information to separate workbook Aspose.Cells | detect external workbook paths in formulas Aspose.Cells | iterate cells and collect external file paths C#

using System;
using System.IO;
using Aspose.Cells;

namespace ExternalLinkReport
{
    // The C# program loads an existing .xlsx workbook, scans every worksheet and cell for formulas, identifies formulas that contain external links by parsing text between '[' and ']', extracts the external file path, and writes the worksheet name, cell address, full formula, and extracted path into a new workbook that serves as a detailed report of all external link formulas.
    class Program
    {
        static void Main(string[] args)
        {
            const string sourcePath = "InputWorkbook.xlsx";
            const string reportPath = "ExternalLinksReport.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            Workbook sourceWorkbook;
            try
            {
                // Load the source workbook
                sourceWorkbook = new Workbook(sourcePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Create a new workbook for the report
            var reportWorkbook = new Workbook();
            var reportSheet = reportWorkbook.Worksheets[0];

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("Cell");
            reportSheet.Cells[0, 2].PutValue("Formula");
            reportSheet.Cells[0, 3].PutValue("External File Path");

            int reportRow = 1; // Start after header

            // Iterate through all worksheets in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Iterate through all cells in the worksheet
                foreach (Cell cell in ws.Cells)
                {
                    if (!cell.IsFormula) continue;

                    string formula = cell.Formula;

                    // Identify external link pattern using '[' and ']'
                    int startIdx = formula.IndexOf('[');
                    int endIdx = formula.IndexOf(']');

                    if (startIdx >= 0 && endIdx > startIdx)
                    {
                        // Extract the external file name from the formula
                        string externalFileName = formula.Substring(startIdx + 1, endIdx - startIdx - 1);
                        string externalPath = externalFileName; // Use extracted name as path

                        // Populate the report row
                        reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(cell.Name);
                        reportSheet.Cells[reportRow, 2].PutValue(formula);
                        reportSheet.Cells[reportRow, 3].PutValue(externalPath);
                        reportRow++;
                    }
                }
            }

            try
            {
                // Save the report workbook
                reportWorkbook.Save(reportPath);
                Console.WriteLine($"Report saved to {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save report: {ex.Message}");
            }
        }
    }
}
