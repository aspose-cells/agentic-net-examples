// Title: Generate an Excel report of formulas that reference external workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, finds all cells whose formulas contain external workbook references (e.g., "[Path]Sheet!A1"), extracts the referenced file path, and creates a new Excel file listing worksheet, cell address, formula, and target path. | Update the example to also parse the external workbook name and sheet name from each formula and add separate columns for these values in the generated report. | Add a command‑line option that lets the user choose between saving the external‑link report as an .xlsx workbook or as a .csv file.
// Common Searches: Aspose.Cells C# list cells with external workbook links and export to a report | how to extract external reference paths from Excel formulas using Aspose.Cells .NET | create a summary of formulas that point to other workbooks with Aspose.Cells | C# generate Excel file showing which cells contain external links in a workbook
// Tags: Aspose.Cells detect external links in formulas | C# create Excel report of external references | scan workbook for external formula links Aspose.Cells | regex parse external file paths Aspose.Cells | export external link summary as CSV Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace ExternalLinkReport
{
    // The program loads a source workbook, iterates through every worksheet and cell to locate formulas that contain external workbook references (identified by a bracketed pattern), extracts the referenced file path using a regular expression, and writes the details—worksheet name, cell address, formula, and target path—into a new Excel workbook that serves as a report.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file that may contain external links
            string inputFile = @"C:\Input\SourceWorkbook.xlsx";

            // Output report file (Excel)
            string reportFile = @"C:\Output\ExternalLinksReport.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(reportFile);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook
                Workbook wb = new Workbook(inputFile);

                // List to hold report rows
                var reportRows = new List<ReportRow>();

                // Regex to extract external file path from a formula
                Regex externalLinkRegex = new Regex(@"\[([^\]]+)\]", RegexOptions.Compiled);

                // Iterate through all worksheets
                foreach (Worksheet sheet in wb.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all cells that contain formulas
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Check if the formula contains an external link pattern
                            Match match = externalLinkRegex.Match(formula);
                            if (match.Success)
                            {
                                // Captured group may include full path
                                string externalReference = match.Groups[1].Value.Trim('\'', '\"');

                                // Add a row to the report
                                reportRows.Add(new ReportRow
                                {
                                    WorksheetName = sheet.Name,
                                    CellName = cell.Name,
                                    Formula = formula,
                                    TargetPath = externalReference
                                });
                            }
                        }
                    }
                }

                // Create a new workbook for the report
                Workbook reportWb = new Workbook();
                Worksheet reportSheet = reportWb.Worksheets[0];
                reportSheet.Name = "ExternalLinksReport";

                // Write header
                Cells reportCells = reportSheet.Cells;
                reportCells["A1"].PutValue("Worksheet");
                reportCells["B1"].PutValue("Cell");
                reportCells["C1"].PutValue("Formula");
                reportCells["D1"].PutValue("Target File Path");

                // Write data rows
                int rowIndex = 1; // zero‑based index; row 1 is the second row in the sheet
                foreach (var row in reportRows)
                {
                    reportCells[rowIndex, 0].PutValue(row.WorksheetName);
                    reportCells[rowIndex, 1].PutValue(row.CellName);
                    reportCells[rowIndex, 2].PutValue(row.Formula);
                    reportCells[rowIndex, 3].PutValue(row.TargetPath);
                    rowIndex++;
                }

                // Auto‑fit columns for better readability
                reportSheet.AutoFitColumns();

                // Save the report
                reportWb.Save(reportFile);
                Console.WriteLine($"Report saved to: {reportFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Simple DTO to hold report information
        private class ReportRow
        {
            public string WorksheetName { get; set; }
            public string CellName { get; set; }
            public string Formula { get; set; }
            public string TargetPath { get; set; }
        }
    }
}
