// Title: Load an Excel workbook from a byte array using Aspose.Cells MemoryStream in C#
// AI Prompts: Write C# code that reads an Excel file into a byte array, creates a MemoryStream, and opens it with Aspose.Cells Workbook. | Show how to process an Excel workbook directly from a MemoryStream without writing any temporary files using Aspose.Cells. | Demonstrate retrieving the name of the first worksheet after loading an Excel file from a byte array with Aspose.Cells.
// Common Searches: C# Aspose.Cells load workbook from byte array memory stream example | How to open Excel file stored in database with Aspose.Cells without saving to disk | Read Excel file into memory and get first sheet name using Aspose.Cells .NET | Aspose.Cells load workbook from stream and access worksheets
// Tags: load workbook from memory stream Aspose.Cells C# | open Excel from byte array Aspose.Cells | process Excel in memory without temporary file | retrieve first worksheet name Aspose.Cells | Aspose.Cells workbook initialization from stream

using System;
using System.IO;
using Aspose.Cells;

// // Reads an Excel file into a byte array, creates a MemoryStream, loads the workbook with Aspose.Cells, and prints the name of the first worksheet.
class Program
{
    static void Main()
    {
        // Load the Excel file bytes from a source (e.g., file, database, network)
        // Replace the path with your actual source of the Excel byte array
        byte[] excelBytes = File.ReadAllBytes("sample.xlsx");

        // Create a memory stream from the byte array
        using (MemoryStream memoryStream = new MemoryStream(excelBytes))
        {
            // Load the workbook from the memory stream
            Workbook workbook = new Workbook(memoryStream);

            // Example processing: output the name of the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine("First worksheet name: " + firstSheet.Name);
        }
    }
}
