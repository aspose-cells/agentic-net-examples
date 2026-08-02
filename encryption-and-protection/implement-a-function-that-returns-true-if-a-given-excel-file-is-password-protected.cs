// Title: Check if an Excel file is password‑protected with Aspose.Cells for .NET (C#)
// Description: A C# helper method uses Aspose.Cells' FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object and returns its IsEncrypted flag, indicating whether the specified Excel workbook requires a password to open.
// Keywords: Aspose.Cells | C# | .NET | Excel password protection | encrypted workbook detection | FileFormatUtil | IsEncrypted property | detect Excel encryption | Excel file security
// Common Searches: Aspose.Cells detect encrypted Excel file C# | How to know if .xlsx is password protected using Aspose | C# check Excel workbook encryption Aspose.Cells | IsExcelFilePasswordProtected example | FileFormatUtil IsEncrypted usage
// Developer Intent: Determine programmatically whether an Excel workbook requires a password before it can be opened.
// Use Cases: Validate uploaded spreadsheets and reject those that are password‑protected. | Skip encrypted files during batch import or conversion pipelines. | Show a warning message to users when they try to open a protected workbook. | Log encrypted files for compliance or audit reporting. | Prompt for a password only after confirming that the file is encrypted.
// AI Prompts: Generate a robust C# method with Aspose.Cells that returns true if an Excel file is encrypted, including exception handling for missing or inaccessible files. | Create unit tests for IsExcelFilePasswordProtected covering encrypted, unencrypted, and non‑existent file scenarios. | Explain how FileFormatUtil.DetectFileFormat and FileFormatInfo.IsEncrypted work together to detect password protection across .xls, .xlsx, and .xlsb formats. | Show how to integrate the password‑check into an ASP.NET file‑upload workflow that blocks encrypted workbooks.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordCheck
{
    // A C# helper method uses Aspose.Cells' FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object and returns its IsEncrypted flag, indicating whether the specified Excel workbook requires a password to open.
    public static class ExcelProtectionHelper
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>True if the file requires a password to open; otherwise, false.</returns>
        public static bool IsExcelFilePasswordProtected(string filePath)
        {
            // Detect the file format and obtain information about the file.
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

            // The IsEncrypted property indicates whether the document is encrypted
            // and therefore requires a password to open.
            return fileInfo.IsEncrypted;
        }

        // Example usage
        public static void Main()
        {
            string path = "sample.xlsx";

            bool protectedFlag = IsExcelFilePasswordProtected(path);
            Console.WriteLine($"Is '{path}' password protected? {protectedFlag}");
        }
    }
}
