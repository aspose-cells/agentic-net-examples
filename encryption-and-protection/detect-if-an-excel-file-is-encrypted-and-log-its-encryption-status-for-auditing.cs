// Title: Check Excel Workbook Encryption and Log Results with Aspose.Cells for .NET
// Description: A C# console utility that leverages Aspose.Cells FileFormatUtil.DetectFileFormat to identify whether an Excel file is encrypted, prints the status, and writes a UTC‑timestamped entry to an audit log for compliance monitoring (US, EU, UK).
// Keywords: Aspose.Cells encryption detection | C# Excel IsEncrypted | FileFormatUtil DetectFileFormat | audit log Excel encryption | password‑protected workbook check | .NET compliance scanning | GDPR Excel security
// Common Searches: how to detect encrypted .xlsx using Aspose.Cells | c# log excel file encryption status | Aspose.Cells FileFormatUtil IsEncrypted example | audit encrypted Excel workbooks .NET | check password protection of Excel files programmatically
// Developer Intent: Determine if a given Excel workbook is encrypted and record the outcome for audit purposes.
// Use Cases: Scheduled compliance job that scans a directory of workbooks and appends encryption results to a central log. | Document management system that blocks upload of password‑protected Excel files. | Security audit script that generates a report of encrypted files from the log entries.
// AI Prompts: Generate C# code to iterate through all .xlsx files in a folder, use Aspose.Cells to check IsEncrypted, and export the findings to a CSV report. | Show how to extend the audit entry with file size, SHA‑256 hash, and user identifier. | Provide error‑handling patterns for unsupported formats or corrupted Excel files when calling FileFormatUtil.DetectFileFormat.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelEncryptionAudit
{
    // A C# console utility that leverages Aspose.Cells FileFormatUtil.DetectFileFormat to identify whether an Excel file is encrypted, prints the status, and writes a UTC‑timestamped entry to an audit log for compliance monitoring (US, EU, UK).
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be checked.
            // If a path is provided as a command‑line argument it will be used,
            // otherwise a default file name is assumed.
            string filePath = args.Length > 0 ? args[0] : "example.xlsx";

            // Detect the file format and retrieve encryption information.
            // This uses the FileFormatUtil.DetectFileFormat(string) rule.
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Determine whether the file is encrypted.
            bool isEncrypted = fileInfo.IsEncrypted;

            // Output the result to the console for immediate feedback.
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Is Encrypted: {isEncrypted}");

            // Append the audit information to a log file.
            string logPath = "encryption_audit_log.txt";
            string logEntry = $"{DateTime.UtcNow:u} | File: {filePath} | Encrypted: {isEncrypted}";
            File.AppendAllText(logPath, logEntry + Environment.NewLine);
        }
    }
}
