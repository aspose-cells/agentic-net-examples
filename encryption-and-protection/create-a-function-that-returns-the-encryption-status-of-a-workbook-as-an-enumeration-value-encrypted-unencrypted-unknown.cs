// Title: Get Excel Workbook Encryption Status with Aspose.Cells for .NET
// Description: C# sample that leverages Aspose.Cells FileFormatUtil.DetectFileFormat to identify if an Excel file is encrypted, unencrypted, or unknown, returning an EncryptionStatus enum and safely handling detection errors.
// Keywords: Aspose.Cells | C# | Excel encryption detection | FileFormatUtil | DetectFileFormat | IsEncrypted | .NET | EncryptionStatus enum | password‑protected workbook
// Common Searches: Aspose.Cells check if Excel file is encrypted C# | How to detect workbook encryption without opening file | Get encryption status of .xlsx using Aspose.Cells | C# enum for Excel file encryption state | FileFormatUtil DetectFileFormat encrypted workbook
// Developer Intent: Retrieve the encryption state of an Excel workbook as an enum value.
// Use Cases: Validate uploaded spreadsheets to reject password‑protected files before processing. | Batch‑scan a directory of workbooks and log each file as Encrypted, Unencrypted, or Unknown. | Integrate a pre‑upload check in a web API that automatically blocks encrypted Excel files.
// AI Prompts: Generate a C# method that uses Aspose.Cells to return an enum indicating whether a workbook is encrypted, unencrypted, or unknown, with exception handling. | Show sample code that iterates over multiple Excel files and prints their encryption status using the GetEncryptionStatus helper. | Explain how FileFormatUtil.DetectFileFormat works and why it avoids loading the entire workbook when checking encryption.

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

    // C# sample that leverages Aspose.Cells FileFormatUtil.DetectFileFormat to identify if an Excel file is encrypted, unencrypted, or unknown, returning an EncryptionStatus enum and safely handling detection errors.
    public static class WorkbookEncryptionHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>EncryptionStatus indicating Encrypted, Unencrypted, or Unknown.</returns>
        public static EncryptionStatus GetEncryptionStatus(string filePath)
        {
            try
            {
                // Detect file format without loading the entire workbook.
                // FileFormatInfo provides the IsEncrypted property.
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

                // If detection succeeded, use the IsEncrypted flag.
                return formatInfo.IsEncrypted ? EncryptionStatus.Encrypted : EncryptionStatus.Unencrypted;
            }
            catch (Exception)
            {
                // Any exception (e.g., file not found, unsupported format) results in Unknown status.
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
            Console.WriteLine($"Workbook encryption status: {status}");
        }
    }
}
