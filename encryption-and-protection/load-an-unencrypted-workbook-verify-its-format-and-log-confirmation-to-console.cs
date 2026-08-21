// Title: Detect Excel file format and encryption with Aspose.Cells in C# before loading
// Description: Demonstrates using Aspose.Cells FileFormatUtil.DetectFileFormat to read an Excel file’s LoadFormat and encryption flag, log the information, and instantiate a Workbook only when the file is not password‑protected.
// Keywords: Aspose.Cells FileFormatUtil | detect Excel format C# | check workbook encryption C# | load unencrypted workbook Aspose | Excel password protection detection | C# Aspose.Cells example | DetectFileFormat usage
// Common Searches: Aspose.Cells detect if Excel file is encrypted | C# check Excel password protection before opening | How to get LoadFormat of Excel file using Aspose | FileFormatUtil DetectFileFormat example | Prevent exception when loading encrypted workbook Aspose.Cells
// Developer Intent: Identify the Excel file type and encryption status, then load the workbook only if it is not password‑protected.
// Use Cases: Validate incoming user‑uploaded spreadsheets to ensure they are unencrypted before processing. | Log file format and worksheet count for audit trails in automated data‑import pipelines. | Avoid runtime errors by pre‑checking encryption status prior to creating a Workbook object.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect an Excel file’s format and encryption flag, then opens it only when it is not password‑protected. | Explain how FileFormatUtil.DetectFileFormat works and why checking IsEncrypted prevents exceptions when loading workbooks. | Provide a sample that prompts for a password and opens an encrypted workbook with Aspose.Cells in C#.

using System;
using Aspose.Cells;

// Demonstrates using Aspose.Cells FileFormatUtil.DetectFileFormat to read an Excel file’s LoadFormat and encryption flag, log the information, and instantiate a Workbook only when the file is not password‑protected.
class Program
{
    static void Main()
    {
        // Path to the workbook file (replace with your actual file path)
        string filePath = "sample.xlsx";

        // Detect the file format and encryption status without loading the workbook
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"Detected LoadFormat: {formatInfo.LoadFormat}");
        Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

        // Verify that the workbook is not encrypted before loading
        if (!formatInfo.IsEncrypted)
        {
            // Load the unencrypted workbook
            Workbook workbook = new Workbook(filePath);
            Console.WriteLine("Workbook loaded successfully.");

            // Additional confirmation (e.g., number of worksheets)
            Console.WriteLine($"Worksheet count: {workbook.Worksheets.Count}");
        }
        else
        {
            Console.WriteLine("The workbook is encrypted and cannot be opened without a password.");
        }
    }
}
