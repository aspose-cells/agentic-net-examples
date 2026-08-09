// Title: Detect Excel Workbook Format and Encryption from a Stream with Aspose.Cells for .NET
// Description: A C# example that opens an Excel file as a read‑only stream, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object, and writes the detected FileFormatType and encryption flag to the console without loading the workbook.
// Keywords: Aspose.Cells | C# | DetectFileFormat | FileFormatUtil | FileFormatInfo | Excel format detection | encrypted workbook | stream processing | read‑only FileStream | Excel file validation
// Common Searches: Aspose.Cells detect Excel format from stream | How to check if an Excel file is encrypted using Aspose.Cells .NET | FileFormatUtil DetectFileFormat example C# | Get workbook file type without opening it Aspose.Cells | Identify Excel file format and encryption status programmatically
// Developer Intent: Determine the workbook’s file format and whether it is encrypted directly from a stream.
// Use Cases: Validate uploaded files are supported Excel formats before further processing. | Log workbook format and encryption status for compliance or audit trails. | Skip or route encrypted workbooks by checking the IsEncrypted flag early in a pipeline.
// AI Prompts: Write C# code that uses Aspose.Cells to detect the format of an Excel file from a MemoryStream and returns the FileFormatType. | Show how to handle encrypted Excel workbooks after detecting them with FileFormatUtil in Aspose.Cells. | Create an ASP.NET Core controller action that validates an uploaded Excel file's format and encryption status using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# example that opens an Excel file as a read‑only stream, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object, and writes the detected FileFormatType and encryption flag to the console without loading the workbook.
    public class DetectWorkbookFormatFromStream
    {
        // Detects the workbook format from a stream and logs the result
        public static void Run(string filePath)
        {
            // Open the file as a read‑only stream
            using (FileStream stream = File.OpenRead(filePath))
            {
                // Use Aspose.Cells utility to detect the format
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream);

                // Output detected format type and encryption status
                Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
            }
        }

        // Simple console entry point for demonstration
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: DetectWorkbookFormatFromStream <path-to-excel-file>");
                return;
            }

            string path = args[0];
            Run(path);
        }
    }
}
