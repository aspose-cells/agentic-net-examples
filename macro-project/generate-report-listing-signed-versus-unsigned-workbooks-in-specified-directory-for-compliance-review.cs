// Title: C# – Create CSV report of signed vs unsigned Excel workbooks with Aspose.Cells
// Description: Scans a given folder for Excel files, loads each workbook via Aspose.Cells, reads the Workbook.IsDigitallySigned flag, and writes a CSV that records the full path and signature status (true, false or error).
// Keywords: Aspose.Cells | C# | .NET | Excel digital signature | IsDigitallySigned | compliance CSV | file enumeration | security audit | workbook signature check | batch Excel processing
// Common Searches: list Excel files with digital signatures using Aspose.Cells | C# generate CSV of signed workbooks | how to check IsDigitallySigned for multiple Excel files | Aspose.Cells batch signature verification | create compliance report for Excel macros
// Developer Intent: Produce a CSV that shows each Excel workbook in a directory and whether it is digitally signed.
// Use Cases: Regulatory audit to confirm all spreadsheets are signed before distribution | Automated build step that flags unsigned macro‑enabled workbooks | Inventory of signed versus unsigned Excel assets for legal record‑keeping
// AI Prompts: Write a function that recursively scans folders for Excel files and returns a map of file paths to their digital‑signature status using Aspose.Cells. | Extend the sample to include a timestamp column and log errors to a separate file. | Generate PowerShell code that invokes the compiled .NET tool and uploads the resulting CSV to SharePoint.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ComplianceReportGenerator
{
    // Scans a given folder for Excel files, loads each workbook via Aspose.Cells, reads the Workbook.IsDigitallySigned flag, and writes a CSV that records the full path and signature status (true, false or error).
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the directory containing the workbooks to scan
            string sourceDirectory = @"C:\Workbooks";

            // Specify the output report file (CSV format)
            string reportPath = Path.Combine(sourceDirectory, "ComplianceReport.csv");

            // Prepare a list to hold report lines
            List<string> reportLines = new List<string>();
            // Add CSV header
            reportLines.Add("FilePath,IsDigitallySigned");

            // Define the file extensions to consider as Excel workbooks
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm" };

            // Enumerate all files with the defined extensions in the directory (non‑recursive)
            foreach (string filePath in Directory.GetFiles(sourceDirectory))
            {
                if (Array.Exists(extensions, ext => ext.Equals(Path.GetExtension(filePath), StringComparison.OrdinalIgnoreCase)))
                {
                    try
                    {
                        // Load the workbook using the constructor that accepts a file path
                        Workbook workbook = new Workbook(filePath);

                        // Determine if the workbook is digitally signed
                        bool isSigned = workbook.IsDigitallySigned;

                        // Add the result to the report
                        reportLines.Add($"{filePath},{isSigned}");
                    }
                    catch (Exception ex)
                    {
                        // If loading fails, record the error state
                        reportLines.Add($"{filePath},Error: {ex.Message}");
                    }
                }
            }

            // Write all report lines to the CSV file
            File.WriteAllLines(reportPath, reportLines);

            Console.WriteLine($"Compliance report generated at: {reportPath}");
        }
    }
}
