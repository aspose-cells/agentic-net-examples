// Title: C# – Detect Encrypted Excel Workbook with Aspose.Cells and Show Warning
// Description: Leverages Aspose.Cells FileFormatUtil.DetectFileFormat to read an Excel file’s format, checks the IsEncrypted flag, writes a console warning for password‑protected files, and loads the workbook only when it is not encrypted.
// Keywords: Aspose.Cells encryption detection | FileFormatUtil IsEncrypted | C# check Excel password protection | detect encrypted workbook .NET | Excel file encryption status | Aspose.Cells FileFormatUtil | encrypted workbook warning | prevent opening encrypted Excel | Aspose.Cells security
// Common Searches: Aspose.Cells how to know if Excel file is password protected | C# detect encrypted .xlsx without opening | FileFormatUtil IsEncrypted example | check Excel workbook encryption Aspose.Cells | display warning for encrypted Excel file C#
// Developer Intent: Identify whether an Excel file is password‑protected before loading it and output a warning if encryption is present.
// Use Cases: Skip encrypted files in batch jobs to avoid runtime errors. | Provide immediate user feedback when a protected workbook is selected. | Validate encryption status prior to processing to ensure smooth workflow.
// AI Prompts: Generate C# code that uses Aspose.Cells to determine if a .xlsx file is encrypted and prints a warning. | Show how to use FileFormatUtil.DetectFileFormat to check encryption and then conditionally open the workbook. | Create a reusable method returning a bool for the encryption status of an Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

// Leverages Aspose.Cells FileFormatUtil.DetectFileFormat to read an Excel file’s format, checks the IsEncrypted flag, writes a console warning for password‑protected files, and loads the workbook only when it is not encrypted.
class WorkbookEncryptionChecker
{
    static void Main()
    {
        // Path to the workbook to be examined
        string filePath = "sample.xlsx";

        // Detect the file format and encryption status without opening the workbook
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

        // Display a warning if the workbook is encrypted
        if (formatInfo.IsEncrypted)
        {
            Console.WriteLine("Warning: The workbook is encrypted and requires a password to open.");
        }
        else
        {
            Console.WriteLine("The workbook is not encrypted.");
        }

        // Optional: If you need to work with the workbook after confirming it is not encrypted,
        // you can load it normally.
        if (!formatInfo.IsEncrypted)
        {
            Workbook workbook = new Workbook(filePath);
            // Perform further operations on the workbook here
        }
    }
}
