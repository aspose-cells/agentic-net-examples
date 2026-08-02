// Title: C# – Generate CSV report of Excel file names, formats and encryption status with Aspose.Cells
// Description: A console utility that scans a given directory, uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify each workbook’s format (XLS, XLSX, CSV, etc.) and encryption flag, then writes a CSV file (FileName,Format,IsEncrypted) to the same folder. Includes basic folder validation and exception handling.
// Keywords: Aspose.Cells detect file format | C# Excel encryption status | CSV inventory of Excel files | FileFormatUtil IsEncrypted | list Excel workbooks folder | bulk Excel file audit | Aspose.Cells .NET example
// Common Searches: how to list Excel files with format and password protection in C# | generate CSV report of Excel workbook types using Aspose.Cells | C# code to check if Excel files are encrypted | Aspose.Cells detect encryption and export to CSV
// Developer Intent: Create a CSV file that enumerates every Excel workbook in a folder, showing its detected format and whether it is encrypted.
// Use Cases: Perform a quick audit of a shared drive to locate password‑protected workbooks before batch processing. | Maintain a compliance inventory that records file type and protection status for all Excel files in a department folder. | Schedule an automated nightly job that updates a CSV summary of newly added or modified Excel files with their encryption flags.
// AI Prompts: Write a C# method that iterates through all files in a folder, uses Aspose.Cells FileFormatUtil.DetectFileFormat to obtain FileFormatType and IsEncrypted, and writes the results to a CSV report. | Add robust error handling to the CSV generation code so that unreadable files are logged and processing continues with the remaining files. | Modify the sample to filter only Excel extensions (xls, xlsx, xlsm, xlsb) before calling DetectFileFormat, and include a command‑line argument for the output CSV path.

using System;
using System.IO;
using Aspose.Cells;

// A console utility that scans a given directory, uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify each workbook’s format (XLS, XLSX, CSV, etc.) and encryption flag, then writes a CSV file (FileName,Format,IsEncrypted) to the same folder. Includes basic folder validation and exception handling.
class ExcelFolderReport
{
    static void Main()
    {
        try
        {
            // Folder containing the Excel files
            string folderPath = @"C:\Path\To\ExcelFolder";

            // Verify that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Output CSV report file
            string reportPath = Path.Combine(folderPath, "ExcelFilesReport.csv");

            // Create or overwrite the report file
            using (var writer = new StreamWriter(reportPath, false))
            {
                // Write CSV header
                writer.WriteLine("FileName,Format,IsEncrypted");

                // Iterate through all files in the folder
                foreach (string filePath in Directory.GetFiles(folderPath))
                {
                    // Detect file format and encryption state
                    FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);

                    // Prepare CSV line
                    string fileName = Path.GetFileName(filePath);
                    string format = info.FileFormatType.ToString();
                    string encrypted = info.IsEncrypted.ToString();

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
