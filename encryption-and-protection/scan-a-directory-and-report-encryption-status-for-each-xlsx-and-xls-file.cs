// Title: C# .NET: Scan a folder for .xls/.xlsx files and report encryption status with Aspose.Cells
// Description: A console utility that enumerates files in a specified directory, filters for Excel workbooks (.xls, .xlsx), uses Aspose.Cells FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object, and prints each file name with its IsEncrypted flag.
// Keywords: Aspose.Cells | C# | .NET | FileFormatUtil | DetectFileFormat | IsEncrypted | Excel encryption detection | password‑protected workbook | list encrypted Excel files | scan directory for XLSX | XLS encryption check
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells C# | C# code to list encrypted .xls and .xlsx files in a folder | detect password‑protected Excel workbooks with Aspose.Cells | scan directory for encrypted Excel files .NET | Aspose.Cells IsEncrypted example
// Developer Intent: Determine which Excel files in a given folder are encrypted.
// Use Cases: Generate a pre‑processing report that flags encrypted workbooks before batch conversion to PDF. | Automatically skip or relocate encrypted spreadsheets during data‑extraction pipelines. | Log encryption status of all Excel files in a directory for compliance or audit purposes.
// AI Prompts: Create a recursive version of this program that writes the file name and encryption flag to a CSV file. | Add error handling to capture and log files that cannot be read or are corrupted while detecting encryption. | Modify the sample to accept multiple folder paths and output a summary count of encrypted vs. unencrypted files.

using System;
using System.IO;
using Aspose.Cells;

// A console utility that enumerates files in a specified directory, filters for Excel workbooks (.xls, .xlsx), uses Aspose.Cells FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object, and prints each file name with its IsEncrypted flag.
class Program
{
    static void Main(string[] args)
    {
        // Use the first argument as the target folder, otherwise use the current directory
        string folderPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

        // Retrieve all files in the folder (non‑recursive)
        string[] allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in allFiles)
        {
            // Process only .xlsx and .xls extensions (case‑insensitive)
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls")
                continue;

            // Detect the file format and obtain encryption information
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Output the file name and its encryption status
            Console.WriteLine($"{Path.GetFileName(filePath)} - Encrypted: {formatInfo.IsEncrypted}");
        }
    }
}
