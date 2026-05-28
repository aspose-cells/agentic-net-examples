using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSummaryReport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // List of workbook file paths to process.
                string[] workbookPaths = new string[]
                {
                    "Sample1.xlsx",
                    "Sample2.xlsm",
                    "Sample3.xls"
                };

                // Create a workbook for the summary report.
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];

                // Write header row.
                reportSheet.Cells["A1"].PutValue("Workbook Path");
                reportSheet.Cells["B1"].PutValue("Check Compatibility");
                reportSheet.Cells["C1"].PutValue("Ribbon Status");

                int currentRow = 1; // Row index for data (zero‑based).

                foreach (string path in workbookPaths)
                {
                    // Verify that the source file exists.
                    if (!File.Exists(path))
                    {
                        reportSheet.Cells[currentRow, 0].PutValue(path);
                        reportSheet.Cells[currentRow, 1].PutValue("N/A");
                        reportSheet.Cells[currentRow, 2].PutValue("File Not Found");
                        currentRow++;
                        continue;
                    }

                    try
                    {
                        // Load the workbook.
                        Workbook wb = new Workbook(path);

                        // Retrieve settings.
                        bool checkCompatibility = wb.Settings.CheckCompatibility;
                        string ribbonStatus = string.IsNullOrEmpty(wb.RibbonXml) ? "Not Set" : "Set";

                        // Write data to the report.
                        reportSheet.Cells[currentRow, 0].PutValue(path);
                        reportSheet.Cells[currentRow, 1].PutValue(checkCompatibility);
                        reportSheet.Cells[currentRow, 2].PutValue(ribbonStatus);
                    }
                    catch (Exception ex)
                    {
                        // Record any error that occurs while processing this workbook.
                        reportSheet.Cells[currentRow, 0].PutValue(path);
                        reportSheet.Cells[currentRow, 1].PutValue("Error");
                        reportSheet.Cells[currentRow, 2].PutValue(ex.Message);
                    }
                    finally
                    {
                        currentRow++;
                    }
                }

                // Save the summary report.
                string reportPath = "WorkbookSummaryReport.xlsx";
                reportWorkbook.Save(reportPath);
                Console.WriteLine($"Summary report generated: {Path.GetFullPath(reportPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}