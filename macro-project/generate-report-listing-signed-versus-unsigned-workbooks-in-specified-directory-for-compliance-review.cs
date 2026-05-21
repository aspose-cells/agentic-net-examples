using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ComplianceReport
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Directory containing the workbooks to scan
                string sourceDirectory = @"C:\Workbooks"; // TODO: change to your directory

                // Output report file path
                string reportPath = @"C:\ComplianceReport\SignedStatusReport.xlsx";

                // Collect workbook file paths (common Excel extensions)
                List<string> workbookFiles = new List<string>();
                string[] extensions = new[] { "*.xlsx", "*.xls", "*.xlsm", "*.xlsb", "*.xlsxml", "*.ods" };
                if (Directory.Exists(sourceDirectory))
                {
                    foreach (string ext in extensions)
                    {
                        workbookFiles.AddRange(Directory.GetFiles(sourceDirectory, ext, SearchOption.AllDirectories));
                    }
                }

                // Prepare data for the report
                List<(string FilePath, bool IsSigned)> results = new List<(string, bool)>();
                foreach (string file in workbookFiles)
                {
                    // Ensure the file exists before attempting to load
                    if (!File.Exists(file))
                        continue;

                    bool signed = false;
                    try
                    {
                        // Load the workbook
                        using (Workbook wb = new Workbook(file))
                        {
                            // Check digital signature status
                            signed = wb.IsDigitallySigned;
                        }
                    }
                    catch (CellsException)
                    {
                        // Password‑protected or otherwise unreadable files are treated as not signed
                        signed = false;
                    }
                    catch (Exception)
                    {
                        // Any other loading issue – skip this file
                        continue;
                    }

                    results.Add((file, signed));
                }

                // Create a new workbook for the report
                using (Workbook reportWorkbook = new Workbook())
                {
                    Worksheet sheet = reportWorkbook.Worksheets[0];

                    // Write header
                    sheet.Cells["A1"].PutValue("File Path");
                    sheet.Cells["B1"].PutValue("Digitally Signed");

                    // Write data rows
                    int rowIndex = 1; // zero‑based index; row 1 is the second row (after header)
                    foreach (var item in results)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(item.FilePath);
                        sheet.Cells[rowIndex, 1].PutValue(item.IsSigned);
                        rowIndex++;
                    }

                    // Ensure the output directory exists
                    string? reportDir = Path.GetDirectoryName(reportPath);
                    if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
                    {
                        Directory.CreateDirectory(reportDir);
                    }

                    // Save the report workbook
                    reportWorkbook.Save(reportPath, SaveFormat.Xlsx);
                }

                Console.WriteLine($"Compliance report generated at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}