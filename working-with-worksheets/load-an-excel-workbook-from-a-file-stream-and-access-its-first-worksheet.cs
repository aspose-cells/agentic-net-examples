// Title: Load an Excel workbook from a FileStream and retrieve the first worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load workbook from a FileStream | how to get the first worksheet name after opening an Excel file with Aspose.Cells | using Aspose.Cells read worksheet at index 0 from a stream | open Excel file with FileStream and access worksheets in Aspose.Cells .NET
// Tags: workbook initialization from stream Aspose.Cells | first worksheet retrieval Aspose.Cells C# | worksheet name extraction Aspose.Cells | excel file stream handling Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example opens an Excel file (input.xlsx) via a FileStream, creates an Aspose.Cells Workbook from that stream, accesses the first worksheet using workbook.Worksheets[0], and prints the worksheet's Name to the console.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Open a file stream for reading the workbook
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            // Load the workbook from the stream
            Workbook workbook = new Workbook(stream);

            // Access the first worksheet (index 0)
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Example: read the name of the first worksheet
            Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
        }
    }
}
