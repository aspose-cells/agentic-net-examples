// Title: Detect Excel Workbook Format with Aspose.Cells for .NET (C#)
// Description: C# sample that verifies a file’s existence, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to identify the workbook type, logs the detected FileFormatType, confirms it with Workbook.FileFormat, and handles errors gracefully.
// Keywords: Aspose.Cells FileFormatUtil | Detect Excel file type C# | Workbook.FileFormat property | Identify spreadsheet format .NET | Excel format detection Aspose | C# Excel file validation
// Common Searches: Aspose.Cells detect workbook format | C# get Excel file type using Aspose | How to identify unknown Excel file format .NET | FileFormatUtil DetectFileFormat example | Check Excel file format before opening Aspose
// Developer Intent: Determine and log the format of an Excel workbook.
// Use Cases: Validate the type of user‑uploaded spreadsheets (XLSX, XLS, CSV, etc.) before processing. | Log the detected format for audit trails in batch import jobs. | Prevent unsupported or malicious files from being opened by checking the format first.
// AI Prompts: Write a C# method that receives a file path, uses Aspose.Cells to detect the spreadsheet format, returns the format enum, and includes robust error handling. | Show how to compare FileFormatUtil.DetectFileFormat output with Workbook.FileFormat and log any discrepancies. | Generate code that logs the detected Excel file type and skips processing if the format is not allowed.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# sample that verifies a file’s existence, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to identify the workbook type, logs the detected FileFormatType, confirms it with Workbook.FileFormat, and handles errors gracefully.
    public class DetectWorkbookFormat
    {
        public static void Run(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Detect the file format using the utility method
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Console.WriteLine($"Detected format (FileFormatUtil): {formatInfo.FileFormatType}");

                // Load the workbook and read its FileFormat property for verification
                using (Workbook workbook = new Workbook(filePath))
                {
                    Console.WriteLine($"Workbook.FileFormat property: {workbook.FileFormat}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            string filePath;

            if (args.Length > 0)
            {
                filePath = args[0];
            }
            else
            {
                Console.Write("Enter the path to the Excel file: ");
                filePath = Console.ReadLine();
            }

            DetectWorkbookFormat.Run(filePath);
        }
    }
}
