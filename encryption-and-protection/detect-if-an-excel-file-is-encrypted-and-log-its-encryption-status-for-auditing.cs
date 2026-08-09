// Title: Detect and Log Excel Workbook Encryption with Aspose.Cells for .NET
// Description: C# sample that uses Aspose.Cells FileFormatUtil to determine if an Excel file is password‑protected (IsEncrypted) and writes the file path and encryption flag to the console, enabling quick audit and compliance checks.
// Keywords: Aspose.Cells | C# | .NET | Excel encryption detection | FileFormatUtil | IsEncrypted | password protected workbook | audit Excel files | compliance reporting | GitHub example | US compliance | EU data protection
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells | C# detect password protected workbook without opening it | Aspose.Cells FileFormatUtil IsEncrypted example | log Excel encryption status for compliance | GitHub Aspose.Cells encryption audit sample
// Developer Intent: Determine whether a given Excel workbook requires a password and record the result for auditing or compliance purposes.
// Use Cases: Run the method across a folder of workbooks to produce a compliance report of encrypted files. | Integrate the check into CI/CD pipelines to reject builds that contain password‑protected Excel documents. | Replace console output with a structured logger (e.g., NLog, Serilog) while preserving the encryption detection logic.
// AI Prompts: Generate C# code that scans a directory of Excel files and logs each file's encryption status using Aspose.Cells. | Show how to swap Console.WriteLine with NLog in the encryption audit example. | Explain exception handling for FileFormatUtil.DetectFileFormat when the file is corrupted or unsupported.

using System;
using Aspose.Cells;

namespace AsposeCellsAudit
{
    // C# sample that uses Aspose.Cells FileFormatUtil to determine if an Excel file is password‑protected (IsEncrypted) and writes the file path and encryption flag to the console, enabling quick audit and compliance checks.
    public class EncryptionAuditor
    {
        // Detects whether the specified Excel file is encrypted and logs the result.
        public static void AuditEncryption(string filePath)
        {
            // Use Aspose.Cells FileFormatUtil to detect file format and encryption status.
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

            // The IsEncrypted property indicates if the file requires a password to open.
            bool isEncrypted = fileInfo.IsEncrypted;

            // Log the encryption status (replace with a proper logging framework if needed).
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Is Encrypted: {isEncrypted}");
        }

        // Example usage
        public static void Main()
        {
            // Replace with the path to the Excel file you want to audit.
            string excelFilePath = "sample.xlsx";

            AuditEncryption(excelFilePath);
        }
    }
}
