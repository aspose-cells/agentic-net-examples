// Title: C# – Generate CSV Report of Excel Files with Format and Encryption Status using Aspose.Cells
// Description: A console app that scans a directory, uses Aspose.Cells FileFormatUtil.DetectFileFormat to determine each workbook’s type (XLS, XLSX, CSV, etc.) and whether it is password‑protected, then writes a CSV file containing the file name, detected format, and encryption flag.
// Keywords: Aspose.Cells C# | FileFormatUtil DetectFileFormat | IsEncrypted Excel | CSV inventory of Excel files | .NET Excel file detection | list encrypted workbooks | batch Excel format audit | password‑protected Excel detection | Excel folder scan C# | Aspose.Cells file format report
// Common Searches: C# Aspose.Cells detect Excel file format | How to list encrypted Excel files in a folder | Generate CSV inventory of workbooks using Aspose.Cells | FileFormatUtil IsEncrypted example C# | Batch scan Excel files for protection status .NET
// Developer Intent: Create a CSV inventory that enumerates every Excel file in a specified folder, showing its detected format and whether the file is encrypted.
// Use Cases: Perform a compliance audit of shared drives by cataloguing workbook types and identifying password‑protected files before migration. | Generate a quick inventory of XLS/XLSX/CSV files for reporting or licensing purposes. | Detect encrypted workbooks so that a downstream process can request passwords or skip protected files.
// AI Prompts: Write C# code that scans a directory, uses Aspose.Cells FileFormatUtil to detect format and encryption, and outputs a CSV report. | Add columns for file size and last‑modified date to the generated CSV. | Implement detailed error logging that records files causing detection exceptions while continuing the scan.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelFolderReportApp
{
    // A console app that scans a directory, uses Aspose.Cells FileFormatUtil.DetectFileFormat to determine each workbook’s type (XLS, XLSX, CSV, etc.) and whether it is password‑protected, then writes a CSV file containing the file name, detected format, and encryption flag.
    class ExcelFolderReport
    {
        static void Main()
        {
            try
            {
                // Specify the folder containing Excel files
                string folderPath = @"C:\Path\To\ExcelFolder";

                // Verify that the source folder exists
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"The folder \"{folderPath}\" does not exist.");
                    return;
                }

                // Path for the generated CSV report
                string reportPath = Path.Combine(folderPath, "ExcelFilesReport.csv");

                // Ensure the directory for the report exists (it will, because it's the same folder)
                // Create the StreamWriter for the CSV file
                using (StreamWriter writer = new StreamWriter(reportPath, false))
                {
                    // Write CSV header
                    writer.WriteLine("FileName,Format,IsEncrypted");

                    // Iterate through all files in the folder
                    foreach (string filePath in Directory.GetFiles(folderPath))
                    {
                        // Skip if the file is not accessible
                        if (!File.Exists(filePath))
                            continue;

                        // Detect file format information
                        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

                        // Skip files that cannot be recognized
                        if (formatInfo.FileFormatType == FileFormatType.Unknown)
                            continue;

                        // Prepare CSV line
                        string fileName = Path.GetFileName(filePath);
                        string format = formatInfo.FileFormatType.ToString();
                        string encrypted = formatInfo.IsEncrypted.ToString();

                        // Write line to CSV
                        writer.WriteLine($"{fileName},{format},{encrypted}");
                    }
                }

                Console.WriteLine($"Report generated at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
