// Title: Detect Excel Workbook Format & Encryption from a Stream with Aspose.Cells for .NET
// Description: A C# console sample that opens an Excel file as a read‑only stream, uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the workbook type and whether it is encrypted, and writes the results to the console without loading the full workbook.
// Keywords: Aspose.Cells | FileFormatUtil | DetectFileFormat | C# | .NET | Excel file format detection | encrypted workbook detection | stream based format check | FileFormatInfo | Excel format type
// Common Searches: Aspose.Cells detect Excel format from stream | check if Excel file is encrypted using Aspose.Cells .NET | how to get file format type without opening workbook | C# detect .xlsx vs .xls with Aspose | FileFormatUtil DetectFileFormat example
// Developer Intent: Determine the workbook’s format and encryption status directly from an input stream.
// Use Cases: Validate uploaded spreadsheets are in a supported Excel format before processing. | Reject or route encrypted workbooks based on the IsEncrypted flag. | Log workbook type for audit trails when handling multiple data sources.
// AI Prompts: Create a reusable method that accepts a Stream and returns FileFormatType and IsEncrypted using Aspose.Cells. | Add comprehensive error handling for DetectFileFormat when the stream is corrupted or not an Excel file. | Write unit tests that verify detection of .xlsx, .xls, .xlsb, and encrypted files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFileFormatDetection
{
    // A C# console sample that opens an Excel file as a read‑only stream, uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the workbook type and whether it is encrypted, and writes the results to the console without loading the full workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (can be passed as a command‑line argument)
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            // Open the file as a read‑only stream
            using (Stream stream = File.OpenRead(filePath))
            {
                // Detect the file format from the stream
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream);

                // Log the detected format type and encryption status
                Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
            }
        }
    }
}
