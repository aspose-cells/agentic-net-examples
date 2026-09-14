// Title: Load an Excel workbook from a byte array using Aspose.Cells Workbook constructor with MemoryStream in C#
// AI Prompts: Write C# code that reads an .xlsx file into a byte[] and creates an Aspose.Cells Workbook from a MemoryStream. | Show how to open an Excel workbook directly from in‑memory bytes with Aspose.Cells without writing a temporary file.
// Common Searches: C# Aspose.Cells read Excel file from byte array memory stream | how to create Workbook from in‑memory byte array using Aspose.Cells | Aspose.Cells open workbook from stream without saving to disk | load xlsx from byte[] with Aspose.Cells C# example
// Tags: Aspose.Cells workbook from MemoryStream C# | load Excel bytes into Aspose.Cells Workbook | open xlsx from byte array without disk write | stream‑based workbook initialization Aspose.Cells | C# read Excel bytes with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Reads an .xlsx file into a byte array, wraps it in a MemoryStream, opens it as an Aspose.Cells Workbook, and prints the first worksheet name.
class Program
{
    static void Main()
    {
        // Load the Excel file into a byte array (replace with your source of bytes)
        byte[] excelBytes = File.ReadAllBytes("input.xlsx");

        // Wrap the byte array in a MemoryStream
        using (MemoryStream ms = new MemoryStream(excelBytes))
        {
            // Open the workbook from the stream using the Stream constructor
            Workbook workbook = new Workbook(ms);

            // Example operation: print the name of the first worksheet
            Console.WriteLine("First worksheet: " + workbook.Worksheets[0].Name);
        }
    }
}
