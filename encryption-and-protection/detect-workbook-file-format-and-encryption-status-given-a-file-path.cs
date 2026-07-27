// Title: Detect Excel workbook format and encryption status using Aspose.Cells in C#
// Description: A console app that receives an Excel file path, uses Aspose.Cells FileFormatUtil.DetectFileFormat to retrieve a FileFormatInfo object, and displays the workbook's format, encryption flag, RMS protection flag, and load format while handling errors gracefully.
// Keywords: Aspose.Cells file format detection | C# detect Excel encryption | FileFormatUtil DetectFileFormat example | identify RMS protected workbook | Excel load format Aspose | check password protection Aspose.Cells | determine workbook type C#
// Common Searches: how to detect excel file type with Aspose.Cells | c# check if xlsx is password protected using Aspose | aspnet detect rms protected workbook | asp.net get load format of excel file | aspnet core file format detection aspose cells
// Developer Intent: Find out the workbook's format and whether it is encrypted or RMS‑protected before loading it.
// Use Cases: Validate incoming spreadsheets for supported formats and password protection. | Log file type and security status for compliance when importing data. | Select an optimal loading strategy based on the detected LoadFormat (e.g., XLSX vs CSV).
// AI Prompts: Write C# code that uses Aspose.Cells to return the file format, IsEncrypted, and IsProtectedByRMS values as JSON. | Create a reusable method that wraps FileFormatUtil.DetectFileFormat with detailed logging and supports batch processing of file paths. | Explain how to open a workbook conditionally when FileFormatInfo.IsEncrypted is false, using Aspose.Cells best practices.

using System;
using Aspose.Cells;

namespace AsposeCellsFileFormatDetection
{
    // A console app that receives an Excel file path, uses Aspose.Cells FileFormatUtil.DetectFileFormat to retrieve a FileFormatInfo object, and displays the workbook's format, encryption flag, RMS protection flag, and load format while handling errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Ensure a file path is provided
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the Excel file as a command‑line argument.");
                return;
            }

            string filePath = args[0];

            try
            {
                // Detect the file format and encryption status using Aspose.Cells utility
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

                // Output the detected information
                Console.WriteLine("File Format Detection Result:");
                Console.WriteLine($"File Path          : {filePath}");
                Console.WriteLine($"Detected Format    : {fileInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted       : {fileInfo.IsEncrypted}");
                Console.WriteLine($"Is Protected By RMS: {fileInfo.IsProtectedByRMS}");
                Console.WriteLine($"Load Format        : {fileInfo.LoadFormat}");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during detection
                Console.WriteLine($"Error detecting file format: {ex.Message}");
            }
        }
    }
}
