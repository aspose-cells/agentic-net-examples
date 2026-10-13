// Title: Read an Excel file from a MemoryStream and get the total number of worksheets with Aspose.Cells for .NET
// AI Prompts: Write C# code that accepts a byte array of an XLSX file, creates a MemoryStream, loads it into an Aspose.Cells Workbook, and returns the sheet total. | Generate a method that takes a Stream containing Excel data and prints how many worksheets it contains without writing the file to disk using Aspose.Cells. | Show how to instantiate a Workbook from a MemoryStream and access the Worksheets.Count property in C#.
// Common Searches: aspocells load workbook from memory stream and count sheets | c# get number of worksheets from excel byte array using Aspose.Cells | how to read worksheet count without saving excel file in .NET | retrieve sheet total from stream with Aspose.Cells library
// Tags: Aspose.Cells load workbook from MemoryStream | C# determine sheet total from Excel stream | count Excel sheets using Aspose.Cells API | read workbook without saving file .NET | access Worksheets.Count via Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX file into a byte array, wraps it in a MemoryStream, creates an Aspose.Cells Workbook from the stream, and outputs the total number of worksheets without persisting the file.
class Program
{
    static void Main()
    {
        // Load the Excel file bytes into a memory stream (replace with your actual source)
        byte[] excelBytes = File.ReadAllBytes("input.xlsx"); // placeholder for source data
        using (MemoryStream memoryStream = new MemoryStream(excelBytes))
        {
            // Create a Workbook object from the memory stream
            Workbook workbook = new Workbook(memoryStream);

            // Retrieve the number of worksheets in the workbook
            int worksheetCount = workbook.Worksheets.Count;

            // Display the worksheet count
            Console.WriteLine($"Worksheet count: {worksheetCount}");
        }
    }
}
