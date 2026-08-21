// Title: C# tool to list signed vs unsigned Excel workbooks in a folder with Aspose.Cells
// Description: Scans a specified directory for Excel files (.xlsx, .xls, .xlsm, .xlsb), uses Aspose.Cells Workbook.IsDigitallySigned to separate signed and unsigned workbooks, and generates an Excel report containing the columns "Signature Status" and "Workbook File Name" saved in the same folder.
// Keywords: Aspose.Cells | C# digital signature | Workbook.IsDigitallySigned | list signed Excel files | Excel compliance report | generate workbook report .NET | scan folder for signed workbooks | Excel file signature audit | C# file system scan | Aspose.Cells example
// Common Searches: Aspose.Cells check if workbook is digitally signed | C# list signed Excel files in a directory | Generate Excel report of signed and unsigned workbooks | How to use Workbook.IsDigitallySigned | Create compliance report for Excel signatures with Aspose | C# scan folder for signed workbooks Aspose.Cells
// Developer Intent: Create an Excel report that identifies which workbooks in a given folder are digitally signed and which are not.
// Use Cases: Regulatory audit of financial models before external distribution | Automated compliance check that flags unsigned workbooks in a project | CI/CD pipeline gate that fails a build when unsigned Excel files are detected | Batch processing of legacy workbooks to enforce signing policies | Integration with a document‑management system to tag files by signature status
// AI Prompts: Write C# code using Aspose.Cells to scan a folder for .xlsx, .xls, .xlsm, .xlsb files, determine each workbook's digital signature via IsDigitallySigned, and output a new Excel file with columns "Signature Status" and "Workbook File Name". | Suggest robust error‑handling strategies for files that cannot be opened as workbooks while generating a signature status report with Aspose.Cells. | Explain how to extend the report to include the file's last modified date and file size alongside its signature status. | Provide performance tips for scanning large directories (thousands of files) with Aspose.Cells. | Show how to localize the report headers for different languages using Aspose.Cells.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace WorkbookSignatureReport
{
    // Scans a specified directory for Excel files (.xlsx, .xls, .xlsm, .xlsb), uses Aspose.Cells Workbook.IsDigitallySigned to separate signed and unsigned workbooks, and generates an Excel report containing the columns "Signature Status" and "Workbook File Name" saved in the same folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the directory containing the workbooks to scan.
            // You can pass the path as a command‑line argument or set it directly here.
            string targetDirectory = args.Length > 0 ? args[0] : @"C:\Workbooks";

            // Collections to hold file names based on signature status.
            List<string> signedFiles = new List<string>();
            List<string> unsignedFiles = new List<string>();

            // Supported Excel extensions.
            string[] extensions = new[] { "*.xlsx", "*.xls", "*.xlsm", "*.xlsb" };

            // Scan each supported file type.
            foreach (string ext in extensions)
            {
                foreach (string filePath in Directory.GetFiles(targetDirectory, ext, SearchOption.TopDirectoryOnly))
                {
                    try
                    {
                        // Load the workbook using the provided constructor rule.
                        Workbook wb = new Workbook(filePath);

                        // Determine digital signature status.
                        if (wb.IsDigitallySigned)
                            signedFiles.Add(Path.GetFileName(filePath));
                        else
                            unsignedFiles.Add(Path.GetFileName(filePath));

                        // Release resources.
                        wb.Dispose();
                    }
                    catch (Exception ex)
                    {
                        // If a file cannot be opened as a workbook, treat it as unsigned and log the issue.
                        Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                        unsignedFiles.Add(Path.GetFileName(filePath) + " (Error)");
                    }
                }
            }

            // Create a new workbook to hold the compliance report.
            Workbook report = new Workbook(); // uses the default constructor rule
            Worksheet sheet = report.Worksheets[0];

            // Write headers.
            sheet.Cells["A1"].PutValue("Signature Status");
            sheet.Cells["B1"].PutValue("Workbook File Name");

            int row = 1; // zero‑based index; row 1 is the second row in the sheet.

            // Populate signed workbook entries.
            foreach (string name in signedFiles)
            {
                sheet.Cells[row, 0].PutValue("Signed");
                sheet.Cells[row, 1].PutValue(name);
                row++;
            }

            // Populate unsigned workbook entries.
            foreach (string name in unsignedFiles)
            {
                sheet.Cells[row, 0].PutValue("Unsigned");
                sheet.Cells[row, 1].PutValue(name);
                row++;
            }

            // Auto‑fit columns for better readability.
            sheet.AutoFitColumns();

            // Save the report using the provided Save method.
            string reportPath = Path.Combine(targetDirectory, "WorkbookSignatureReport.xlsx");
            report.Save(reportPath); // uses Save(string) rule

            Console.WriteLine($"Compliance report generated at: {reportPath}");
        }
    }
}
