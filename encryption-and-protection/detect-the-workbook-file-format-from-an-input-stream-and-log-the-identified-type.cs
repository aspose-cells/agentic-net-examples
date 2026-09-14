// Title: Detect Excel workbook format from a Stream using Aspose.Cells for .NET and log it to console
// AI Prompts: Write C# code that reads an Excel file from a Stream, uses Aspose.Cells to auto‑detect its format, and prints the FileFormatType to the console. | Show how to load a workbook from a MemoryStream with Aspose.Cells and retrieve the detected file format without specifying the extension. | Create a snippet that opens a workbook from a FileStream, accesses the Workbook.FileFormat property, and logs the format name for debugging.
// Common Searches: asp.net core detect excel file format from a stream using Aspose.Cells | c# Aspose.Cells get workbook file type after loading from FileStream | how to print detected workbook format with Aspose.Cells .NET library | auto detect xls vs xlsx from input stream Aspose.Cells example
// Tags: auto-detect workbook format Aspose.Cells | Workbook.FileFormat property usage | load workbook from stream C# Aspose.Cells | log detected Excel format console | identify XLSX vs XLS with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Opens an Excel file via a Stream, lets Aspose.Cells automatically determine the workbook format, reads the Workbook.FileFormat property, and writes the detected format (e.g., XLS, XLSX) to the console.
class Program
{
    static void Main()
    {
        // Obtain the input stream (replace with your actual source)
        using (FileStream inputStream = new FileStream("input.xlsx", FileMode.Open, FileAccess.Read))
        {
            // Load the workbook from the stream; Aspose.Cells auto‑detects the format
            Workbook workbook = new Workbook(inputStream);

            // Retrieve the detected file format
            FileFormatType detectedFormat = workbook.FileFormat;

            // Log the identified format
            Console.WriteLine($"Detected workbook format: {detectedFormat}");
        }
    }
}
