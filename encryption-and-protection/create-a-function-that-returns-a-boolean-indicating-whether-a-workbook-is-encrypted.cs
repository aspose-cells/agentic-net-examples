// Title: Determine if an Excel workbook is encrypted with Aspose.Cells for .NET (C#) without loading the file
// Description: Provides a C# static helper that checks a workbook file path, uses Aspose.Cells FileFormatUtil.DetectFileFormat to obtain FileFormatInfo, and returns the IsEncrypted flag. Includes file‑existence validation and clear exception handling.
// Keywords: Aspose.Cells encryption detection | C# workbook encrypted check | FileFormatUtil DetectFileFormat | IsEncrypted property | password protected Excel file | detect encrypted .xlsx | detect encrypted .xls | Aspose.Cells API | .NET Excel security | Excel file encryption status
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells C# | Aspose.Cells detect password protected workbook without opening | C# IsWorkbookEncrypted example | FileFormatUtil DetectFileFormat encrypted flag | determine Excel file encryption status Aspose.Cells
// Developer Intent: Find a quick way to know whether a given Excel workbook is encrypted without fully loading it into memory.
// Use Cases: Skip encrypted files during bulk import or migration processes. | Display an encryption icon next to Excel files in a file‑management UI. | Prevent automated reports from failing by pre‑checking password‑protected workbooks.
// AI Prompts: Write a C# method that uses Aspose.Cells to return true if a workbook is encrypted, handling missing files and wrapping exceptions. | Generate unit tests for IsWorkbookEncrypted covering encrypted, unencrypted, and non‑existent files. | Explain how FileFormatUtil.DetectFileFormat retrieves encryption information without loading the workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Provides a C# static helper that checks a workbook file path, uses Aspose.Cells FileFormatUtil.DetectFileFormat to obtain FileFormatInfo, and returns the IsEncrypted flag. Includes file‑existence validation and clear exception handling.
    public static class WorkbookEncryptionHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>True if the workbook is encrypted; otherwise, false.</returns>
        public static bool IsWorkbookEncrypted(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            try
            {
                // Detect the file format and retrieve encryption information without loading the workbook.
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
                return fileInfo.IsEncrypted;
            }
            catch (Exception ex)
            {
                // Wrap any Aspose.Cells exceptions for clearer error handling.
                throw new InvalidOperationException("Failed to detect workbook encryption.", ex);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Example usage:
                string workbookPath = "encrypted.xlsx";
                bool encrypted = WorkbookEncryptionHelper.IsWorkbookEncrypted(workbookPath);
                Console.WriteLine($"Workbook encrypted: {encrypted}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
