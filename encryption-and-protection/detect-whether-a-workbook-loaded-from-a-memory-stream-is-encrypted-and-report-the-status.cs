// Title: Detect Encrypted Excel Workbook from a MemoryStream with Aspose.Cells for .NET
// Description: Loads an Excel file into a MemoryStream, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to read format information, and reports the IsEncrypted flag to indicate whether the workbook is password‑protected.
// Keywords: Aspose.Cells | C# detect encrypted workbook | FileFormatUtil IsEncrypted | Excel encryption detection .NET | MemoryStream Excel format check | password protected Excel file | Aspose.Cells file format detection
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells | detect password protection in .xlsx with C# | FileFormatUtil DetectFileFormat encrypted workbook example | read Excel file from MemoryStream and verify encryption | Aspose.Cells IsEncrypted property usage
// Developer Intent: Identify whether a workbook loaded from a MemoryStream is encrypted before processing.
// Use Cases: Validate user‑uploaded spreadsheets and prompt for a password only when needed. | Skip decryption steps for plain files to improve processing speed. | Log encryption status for compliance and audit trails in enterprise workflows.
// AI Prompts: Generate C# code that reads an Excel file into a MemoryStream, calls Aspose.Cells.FileFormatUtil.DetectFileFormat, and returns a boolean indicating encryption. | Create a try‑catch example that throws a custom exception when DetectFileFormat reports the workbook is encrypted. | Write unit tests that verify DetectFileFormat.IsEncrypted returns true for a password‑protected .xlsx and false for an unprotected file.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file into a MemoryStream, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to read format information, and reports the IsEncrypted flag to indicate whether the workbook is password‑protected.
class DetectEncryption
{
    public static void Run()
    {
        string filePath = "sample.xlsx";

        // Verify that the file exists before attempting to read it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the Excel file into a memory stream
            using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(filePath)))
            {
                // Detect format information directly from the stream
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(ms);

                // Output whether the workbook is encrypted
                Console.WriteLine($"Workbook is encrypted: {formatInfo.IsEncrypted}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Entry point for the application
    static void Main(string[] args)
    {
        Run();
    }
}
