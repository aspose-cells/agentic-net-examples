// Title: Check if an Excel workbook is encrypted with Aspose.Cells for .NET (C#)
// Description: C# helper that uses Aspose.Cells FileFormatUtil.DetectFileFormat to determine whether a workbook file is password‑protected without loading the document, with file‑existence verification and robust error handling.
// Keywords: Aspose.Cells | C# encryption detection | Excel workbook encrypted | FileFormatUtil IsEncrypted | DetectFileFormat | password protected Excel | WorkbookEncryptionHelper | Aspose.Cells .NET | encrypted workbook check | Excel file security
// Common Searches: Aspose.Cells how to know if Excel file is encrypted | C# detect password protected workbook without opening | FileFormatUtil IsEncrypted example | Check Excel encryption status Aspose | Determine if .xlsx is encrypted using Aspose.Cells
// Developer Intent: Determine whether a given Excel file is encrypted.
// Use Cases: Prompt the user for a password only when the file is encrypted | Skip encrypted workbooks during bulk import or migration | Generate a compliance report listing encryption status of spreadsheets | Validate files before applying transformations or calculations | Automate pre‑processing of spreadsheets in a CI/CD pipeline
// AI Prompts: Generate a C# method using Aspose.Cells that returns true if a workbook is encrypted and logs errors. | Create a script that scans a directory and prints encryption status for each Excel file using the IsWorkbookEncrypted helper. | Explain how to extend IsWorkbookEncrypted to also retrieve the encryption algorithm name from FileFormatInfo. | Provide unit tests for the IsWorkbookEncrypted function covering missing file, unencrypted, and encrypted scenarios. | Show how to integrate IsWorkbookEncrypted into an ASP.NET Core file‑upload workflow.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // C# helper that uses Aspose.Cells FileFormatUtil.DetectFileFormat to determine whether a workbook file is password‑protected without loading the document, with file‑existence verification and robust error handling.
    public static class WorkbookEncryptionHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>True if the workbook is encrypted; otherwise, false.</returns>
        public static bool IsWorkbookEncrypted(string filePath)
        {
            // Ensure the file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return false;
            }

            try
            {
                // Detect the file format and retrieve encryption information without fully loading the workbook.
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                return formatInfo.IsEncrypted;
            }
            catch (Exception ex)
            {
                // Log unexpected errors and treat the workbook as not encrypted.
                Console.WriteLine($"Error detecting encryption for '{filePath}': {ex.Message}");
                return false;
            }
        }

        // Example usage
        public static void Demo()
        {
            string encryptedPath = "encrypted.xlsx";
            string normalPath = "normal.xlsx";

            bool isEncrypted1 = IsWorkbookEncrypted(encryptedPath);
            Console.WriteLine($"'{encryptedPath}' encrypted? {isEncrypted1}");

            bool isEncrypted2 = IsWorkbookEncrypted(normalPath);
            Console.WriteLine($"'{normalPath}' encrypted? {isEncrypted2}");
        }
    }

    // Entry point for the console application.
    public static class Program
    {
        public static void Main()
        {
            try
            {
                WorkbookEncryptionHelper.Demo();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
