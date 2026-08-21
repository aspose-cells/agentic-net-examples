// Title: Detect Excel workbook format and encryption status with Aspose.Cells (C#)
// Description: C# console example that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to identify a spreadsheet’s format (XLS, XLSX, CSV, etc.), determine if it is password‑encrypted or RMS‑protected, and obtain the LoadFormat needed for further processing.
// Keywords: Aspose.Cells file format detection | C# detect Excel encryption | FileFormatUtil DetectFileFormat | FileFormatInfo IsEncrypted | RMS protected spreadsheet | load format Aspose.Cells | identify workbook type .NET | check password protection Excel C#
// Common Searches: How to check if an Excel file is password protected with Aspose.Cells | C# get spreadsheet file type using Aspose.Cells | Determine if a workbook is RMS protected in .NET | Retrieve LoadFormat of an unknown Excel file | Detect encrypted Excel workbook programmatically
// Developer Intent: Find the workbook’s type and whether it is encrypted or RMS‑protected.
// Use Cases: Validate incoming spreadsheets before processing to ensure they are in a supported format and not locked. | Log format and protection details of uploaded workbooks for audit, compliance, or troubleshooting. | Conditionally open a workbook only when IsEncrypted is false; otherwise prompt the user for a password.
// AI Prompts: Generate C# code that opens a workbook with Aspose.Cells only after confirming it is not encrypted, otherwise asks the user for a password. | Write a method that returns a one‑line summary of FileFormatInfo (format, encryption, RMS, load format) for any file path. | Provide robust error‑handling examples for FileFormatUtil.DetectFileFormat when the file is missing, corrupted, or unsupported.

using System;
using Aspose.Cells;

// C# console example that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to identify a spreadsheet’s format (XLS, XLSX, CSV, etc.), determine if it is password‑encrypted or RMS‑protected, and obtain the LoadFormat needed for further processing.
class Program
{
    // Entry point of the console application
    static void Main(string[] args)
    {
        // Get the file path from command‑line arguments or use a default one
        string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

        // Detect the file format and encryption status using Aspose.Cells utility
        FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

        // Output the detection results
        Console.WriteLine($"File Path: {filePath}");
        Console.WriteLine($"Detected File Format Type: {fileInfo.FileFormatType}");
        Console.WriteLine($"Is Encrypted: {fileInfo.IsEncrypted}");
        Console.WriteLine($"Is Protected By RMS: {fileInfo.IsProtectedByRMS}");
        Console.WriteLine($"Load Format: {fileInfo.LoadFormat}");
    }
}
