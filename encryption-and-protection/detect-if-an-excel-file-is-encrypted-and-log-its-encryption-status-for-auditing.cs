// Title: Detect Excel Encryption and Write Audit Log with Aspose.Cells for .NET
// Description: A C# console app that uses Aspose.Cells.FileFormatUtil.DetectFileFormat to check if an Excel workbook is encrypted (IsEncrypted flag). It creates a UTC timestamped entry with the file path and encryption status, prints it to the console, and appends it to an audit.log file for compliance tracking.
// Keywords: Aspose.Cells encryption detection | C# detect encrypted Excel workbook | FileFormatUtil IsEncrypted | audit log Excel encryption | Excel password protection check .NET | compliance audit Excel files | timestamped log entry C# | detect encrypted .xlsx Aspose
// Common Searches: how to check if an Excel file is password protected using Aspose.Cells | C# code to log Excel encryption status for compliance | Aspose.Cells DetectFileFormat encrypted workbook example | write audit log for encrypted Excel files in .NET | detect encrypted .xlsx files programmatically
// Developer Intent: Determine whether a given Excel workbook is encrypted and record the result in a persistent audit log.
// Use Cases: Validate uploaded Excel files on a web server and reject or flag encrypted workbooks before processing. | Generate daily GDPR or SOX compliance reports that list password‑protected spreadsheets in a repository. | Integrate encryption checks into CI/CD pipelines to prevent protected files from being committed to source control. | Maintain an immutable audit trail of encryption status for financial or legal Excel documents.
// AI Prompts: Create a C# method that receives a file path, returns true if the workbook is encrypted using Aspose.Cells, and writes a UTC timestamped entry to a configurable log file. | Show how to wrap FileFormatUtil.DetectFileFormat in try‑catch blocks to handle missing files, unsupported formats, and permission errors while still logging each attempt. | Extend the audit entry to include file size, operating system user name, and optional custom tags for regulatory reporting.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelEncryptionAudit
{
    // A C# console app that uses Aspose.Cells.FileFormatUtil.DetectFileFormat to check if an Excel workbook is encrypted (IsEncrypted flag). It creates a UTC timestamped entry with the file path and encryption status, prints it to the console, and appends it to an audit.log file for compliance tracking.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be checked.
            // If a path is provided as a command‑line argument it is used,
            // otherwise a default file name is assumed.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            // Detect the file format and retrieve encryption information.
            // This uses the FileFormatUtil.DetectFileFormat(string) rule.
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
            bool isEncrypted = fileInfo.IsEncrypted;

            // Build a log entry that includes a timestamp, file name and encryption status.
            string logEntry = $"{DateTime.UtcNow:u} | File: {filePath} | Encrypted: {isEncrypted}";

            // Output the result to the console for immediate feedback.
            Console.WriteLine(logEntry);

            // Append the audit information to a persistent log file.
            const string auditLogPath = "audit.log";
            using (StreamWriter writer = new StreamWriter(auditLogPath, append: true))
            {
                writer.WriteLine(logEntry);
            }
        }
    }
}
