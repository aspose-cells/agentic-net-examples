// Title: Check if an Excel workbook is encrypted with Aspose.Cells in C#
// Description: C# helper method that validates a file path, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object, and returns the IsEncrypted flag as a boolean. Handles missing files and errors without loading the workbook.
// Keywords: Aspose.Cells | C# | detect encrypted Excel file | FileFormatUtil | IsEncrypted | password protected workbook | Excel encryption status | DetectFileFormat | WorkbookEncryptionHelper | .NET Excel security
// Common Searches: Aspose.Cells check if Excel file is encrypted | C# detect password protection on .xlsx using Aspose | How to know if an Excel workbook is encrypted without opening it | FileFormatUtil IsEncrypted example | Determine encryption status of Excel file in .NET
// Developer Intent: Find out whether a given Excel workbook is encrypted or password‑protected using Aspose.Cells.
// Use Cases: Skip encrypted files during bulk import to avoid load errors. | Log encryption status for audit trails in document management systems. | Display a security indicator in a UI that lists Excel documents. | Validate files before applying transformations that require an unprotected workbook.
// AI Prompts: Generate unit tests for WorkbookEncryptionHelper.IsWorkbookEncrypted covering missing files, unencrypted workbooks, and encrypted workbooks. | Provide an alternative approach that attempts to load the workbook with LoadOptions and catches the encryption exception to infer encryption status. | Create a step‑by‑step guide showing how to integrate IsWorkbookEncrypted into a file‑processing pipeline that filters protected Excel files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // C# helper method that validates a file path, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object, and returns the IsEncrypted flag as a boolean. Handles missing files and errors without loading the workbook.
    public static class WorkbookEncryptionHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>True if encrypted; otherwise false.</returns>
        public static bool IsWorkbookEncrypted(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Workbook file not found.", filePath);

            try
            {
                // Detect file format and encryption status without loading the workbook.
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                return formatInfo.IsEncrypted;
            }
            catch (Exception ex)
            {
                // Log the error and return false indicating unknown encryption status.
                Console.Error.WriteLine($"Error detecting encryption: {ex.Message}");
                return false;
            }
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            // Example file path; adjust as needed.
            string filePath = @"C:\Docs\sample.xlsx";

            try
            {
                bool encrypted = WorkbookEncryptionHelper.IsWorkbookEncrypted(filePath);
                Console.WriteLine($"Workbook encrypted: {encrypted}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to check workbook encryption: {ex.Message}");
            }
        }
    }
}
