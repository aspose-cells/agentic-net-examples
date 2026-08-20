// Title: C# .NET Console App to Detect Excel Workbook Encryption with Aspose.Cells
// Description: A simple console program that receives an Excel file path, verifies the file's existence, uses Aspose.Cells FileFormatUtil.DetectFileFormat to obtain FileFormatInfo, checks the IsEncrypted flag, and prints whether the workbook is password‑protected along with a clear user message.
// Keywords: Aspose.Cells encryption detection | C# check Excel password protection | FileFormatUtil IsEncrypted example | detect encrypted workbook .NET | console app Excel encryption status
// Common Searches: How to find out if an Excel file is password protected using Aspose.Cells C# | C# console program to read encryption flag of a workbook | Aspose.Cells FileFormatUtil detect encrypted workbook | Check Excel file encryption status in .NET
// Developer Intent: Identify whether a supplied Excel file is encrypted and inform the user if a password is required.
// Use Cases: Screen user‑uploaded spreadsheets before processing to reject password‑protected files. | Run a nightly batch job that flags encrypted workbooks for manual review. | Add a pre‑deployment validation step that ensures only unencrypted Excel assets are included in a release.
// AI Prompts: Generate a C# method that takes a file path and returns true if the workbook is encrypted using Aspose.Cells. | Create comprehensive error handling for missing files, unsupported formats, and permission issues when checking encryption. | Write unit tests that mock encrypted and unencrypted Excel files to verify the detection logic.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDetector
{
    // A simple console program that receives an Excel file path, verifies the file's existence, uses Aspose.Cells FileFormatUtil.DetectFileFormat to obtain FileFormatInfo, checks the IsEncrypted flag, and prints whether the workbook is password‑protected along with a clear user message.
    class Program
    {
        static void Main(string[] args)
        {
            // Ensure a file path argument is provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsEncryptionDetector <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Verify that the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Detect the file format and retrieve encryption information
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Output encryption status
            Console.WriteLine($"Is file encrypted? {fileInfo.IsEncrypted}");

            // Provide a clear password requirement message
            if (fileInfo.IsEncrypted)
            {
                Console.WriteLine("A password is required to open this workbook.");
            }
            else
            {
                Console.WriteLine("No password is required to open this workbook.");
            }
        }
    }
}
