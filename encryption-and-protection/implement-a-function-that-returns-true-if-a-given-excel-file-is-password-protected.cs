// Title: C# – Detect Password‑Protected Excel Workbook Using Aspose.Cells
// Description: Learn how to use Aspose.Cells for .NET to check if an Excel file (.xls, .xlsx, .xlsb, etc.) is encrypted. The sample code validates the file path, calls FileFormatUtil.DetectFileFormat, and reads FileFormatInfo.IsEncrypted, with robust error handling.
// Keywords: Aspose.Cells C# detect encrypted Excel | check Excel password protection .NET | FileFormatUtil IsEncrypted | Excel file encryption detection | C# verify workbook password | Aspose.Cells file format detection | Excel security audit C# | detect protected .xlsx programmatically
// Common Searches: how to know if an Excel file is password protected in C# | Aspose.Cells detect encrypted workbook | C# check if .xlsx requires a password | FileFormatUtil DetectFileFormat password | Aspose.Cells IsEncrypted example
// Developer Intent: Identify whether a specified Excel workbook is encrypted and requires a password before opening or processing it.
// Use Cases: Validate user‑uploaded spreadsheets to reject or prompt for passwords before import. | Run a security scan over a repository of workbooks and generate a report of encrypted files. | Skip password‑protected files in bulk conversion or data‑extraction pipelines.
// AI Prompts: Generate a C# method that returns true if an Excel file is password protected using Aspose.Cells, including file‑existence checks and exception handling. | Show sample code that iterates through a folder of Excel files and logs each file’s encryption status with Aspose.Cells. | Explain the behavior of FileFormatInfo.IsEncrypted and which Excel formats (xls, xlsx, xlsb, csv, etc.) it can evaluate.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Learn how to use Aspose.Cells for .NET to check if an Excel file (.xls, .xlsx, .xlsb, etc.) is encrypted. The sample code validates the file path, calls FileFormatUtil.DetectFileFormat, and reads FileFormatInfo.IsEncrypted, with robust error handling.
    public static class ExcelProtectionHelper
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>True if the file is encrypted and requires a password to open; otherwise, false.</returns>
        public static bool IsExcelFilePasswordProtected(string filePath)
        {
            // Verify that the file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return false;
            }

            try
            {
                // Detect the file format and retrieve its metadata.
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

                // The IsEncrypted property indicates whether the document is encrypted
                // and therefore requires a password to open.
                return fileInfo.IsEncrypted;
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully.
                Console.WriteLine($"Error while checking protection: {ex.Message}");
                return false;
            }
        }

        // Example usage
        public static void Run()
        {
            string path = "sample.xlsx";

            bool protectedStatus = IsExcelFilePasswordProtected(path);
            Console.WriteLine($"Is '{path}' password protected? {protectedStatus}");
        }
    }

    // Entry point required for console application.
    public static class Program
    {
        public static void Main()
        {
            try
            {
                ExcelProtectionHelper.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
