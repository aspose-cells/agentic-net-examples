// Title: Detect Workbook File Format and Encryption with Aspose.Cells for .NET (C#)
// Description: A C# console sample that iterates over spreadsheet files, uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the file type, encryption status and load format, then loads each workbook to confirm the Workbook.FileFormat property, logging all details to the console.
// Keywords: Aspose.Cells | C# file format detection | FileFormatUtil | FileFormatInfo | detect Excel format | check workbook encryption | Workbook.FileFormat property | identify unknown spreadsheet | Aspose.Cells .NET | load format detection
// Common Searches: Aspose.Cells detect workbook format C# | How to check if Excel file is encrypted using Aspose.Cells | FileFormatUtil DetectFileFormat example | Get file format type before loading workbook Aspose | Identify CSV vs XLSX with Aspose.Cells | Batch detect spreadsheet formats .NET
// Developer Intent: Determine the exact format and encryption state of each spreadsheet file and log the information for downstream processing.
// Use Cases: Process a folder of mixed‑type spreadsheets (XLS, XLSX, CSV, etc.) and record their formats before applying business logic. | Inspect the IsEncrypted flag to decide whether a file needs decryption or should be skipped. | Validate that the Workbook.FileFormat property matches the detected format to avoid loading errors. | Generate audit logs of incoming files for compliance or data‑governance pipelines.
// AI Prompts: Write a C# method that receives a file path and returns a formatted string with the detected file format, encryption status, and load format using Aspose.Cells. | Create code to scan all files in a directory, use FileFormatUtil.DetectFileFormat for each, log the results, and skip files that are encrypted. | Show how to handle unsupported or unknown file formats gracefully when calling FileFormatUtil.DetectFileFormat in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// A C# console sample that iterates over spreadsheet files, uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the file type, encryption status and load format, then loads each workbook to confirm the Workbook.FileFormat property, logging all details to the console.
class DetectWorkbookFormat
{
    static void Main(string[] args)
    {
        // List of workbook files whose format needs to be identified.
        // Replace these with actual file paths as needed.
        string[] files = { "unknown1.xlsx", "unknown2.xls", "unknown3.csv" };

        foreach (string filePath in files)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            // Detect the file format using Aspose.Cells utility.
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Detected FileFormatType : {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted           : {formatInfo.IsEncrypted}");
            Console.WriteLine($"Detected LoadFormat    : {formatInfo.LoadFormat}");

            // Load the workbook to demonstrate the Workbook.FileFormat property.
            using (Workbook wb = new Workbook(filePath))
            {
                Console.WriteLine($"Workbook.FileFormat property : {wb.FileFormat}");
            }

            Console.WriteLine(new string('-', 40));
        }
    }
}
