// Title: Generate an Excel report of formulas with external links and extract their file paths using Aspose.Cells for .NET (C#)
// Common Searches: how to create a report of Excel formulas that reference external workbooks with Aspose.Cells C# | Aspose.Cells C# find formulas containing external file paths | extract external workbook links from Excel formulas using Aspose.Cells | list cells with external references and save to a new workbook in C# | scan all worksheets for external link formulas and export details with Aspose.Cells | C# code to detect [FilePath] pattern in Excel formula using Aspose.Cells
// Tags: Aspose.Cells extract external link formulas | C# generate external link report workbook | scan worksheets for external references Aspose.Cells | detect external file path in Excel formula C# | write formula audit to new Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a workbook, iterates through every worksheet and cell, identifies formulas that contain the '[' and ']' pattern indicating an external link, extracts the file path inside the brackets, and writes the worksheet name, cell address, full formula, and extracted path to a new Excel file that serves as a detailed external‑link report.
class ExternalLinksReporter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string reportPath = "ExternalLinksReport.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook (no password supplied)
            var sourceWorkbook = new Workbook(inputPath);

            // Create a new workbook for the report
            var reportWorkbook = new Workbook();
            var reportSheet = reportWorkbook.Worksheets[0];

            // Write header row
            reportSheet.Cells["A1"].PutValue("Worksheet");
            reportSheet.Cells["B1"].PutValue("Cell");
            reportSheet.Cells["C1"].PutValue("Formula");
            reportSheet.Cells["D1"].PutValue("External File Path");

            int reportRow = 1; // zero‑based index for the second row

            // Scan all worksheets and cells for external links
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                foreach (Cell cell in ws.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Detect external link pattern [FilePath]Sheet!A1
                        if (formula.Contains("[") && formula.Contains("]"))
                        {
                            int startIdx = formula.IndexOf('[') + 1;
                            int endIdx = formula.IndexOf(']', startIdx);
                            if (endIdx > startIdx)
                            {
                                string externalPath = formula.Substring(startIdx, endIdx - startIdx);

                                // Populate report row
                                reportSheet.Cells[reportRow, 0].PutValue(ws.Name);          // Worksheet name
                                reportSheet.Cells[reportRow, 1].PutValue(cell.Name);      // Cell address
                                reportSheet.Cells[reportRow, 2].PutValue(formula);       // Full formula
                                reportSheet.Cells[reportRow, 3].PutValue(externalPath);  // Extracted path

                                reportRow++;
                            }
                        }
                    }
                }
            }

            // Save the report
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Report saved to \"{reportPath}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
