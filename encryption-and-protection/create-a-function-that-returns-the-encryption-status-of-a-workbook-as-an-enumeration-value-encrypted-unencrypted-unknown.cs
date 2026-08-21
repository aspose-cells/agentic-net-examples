// Title: Check Excel Workbook Encryption with Aspose.Cells for .NET
// Description: C# example that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to determine whether an Excel file is encrypted, unencrypted, or unknown, returning the result as an EncryptionStatus enum.
// Keywords: Aspose.Cells | .NET | C# | Excel encryption detection | FileFormatUtil | IsEncrypted | WorkbookEncryptionHelper | EncryptionStatus enum | detect password‑protected workbook | check Excel file protection
// Common Searches: Aspose.Cells check if Excel file is password protected | C# detect encrypted workbook without opening | Get encryption status of .xlsx using Aspose | FileFormatUtil DetectFileFormat encryption | How to handle missing Excel file when checking encryption
// Developer Intent: Implement a method that returns an enumeration (Encrypted, Unencrypted, Unknown) indicating the encryption state of a workbook.
// Use Cases: Filter out password‑protected files before bulk import | Display encryption status in a document‑management dashboard | Log files that cannot be read due to unknown format or missing path | Validate workbook security policy in automated pipelines
// AI Prompts: Generate a C# function using Aspose.Cells that returns EncryptionStatus (Encrypted, Unencrypted, Unknown) for a given file path, with exception handling. | Show how to loop through multiple Excel files, call GetEncryptionStatus, and write results to a CSV log. | Explain the role of FileFormatUtil.DetectFileFormat and its IsEncrypted property in determining workbook protection without loading the file.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionStatus
{
    // Enumeration representing possible encryption states of a workbook
    public enum EncryptionStatus
    {
        Encrypted,
        Unencrypted,
        Unknown
    }

    // C# example that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to determine whether an Excel file is encrypted, unencrypted, or unknown, returning the result as an EncryptionStatus enum.
    public static class WorkbookEncryptionHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>EncryptionStatus indicating whether the workbook is encrypted, unencrypted, or unknown.</returns>
        public static EncryptionStatus GetEncryptionStatus(string filePath)
        {
            try
            {
                // Detect the file format and retrieve encryption information without loading the workbook
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

                // If the format detection succeeded, use the IsEncrypted property
                return formatInfo.IsEncrypted ? EncryptionStatus.Encrypted : EncryptionStatus.Unencrypted;
            }
            catch (Exception)
            {
                // Any exception (e.g., file not found, unsupported format) results in an unknown status
                return EncryptionStatus.Unknown;
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string path = "example.xlsx";

            EncryptionStatus status = WorkbookEncryptionHelper.GetEncryptionStatus(path);
            Console.WriteLine($"Encryption status of '{path}': {status}");
        }
    }
}
