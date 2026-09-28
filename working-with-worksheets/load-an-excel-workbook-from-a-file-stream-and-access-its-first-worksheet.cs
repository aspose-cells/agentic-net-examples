// Title: Load an Excel workbook from a FileStream and retrieve the first worksheet using Aspose.Cells for .NET
// AI Prompts: Read an Excel file from a FileStream, create a Workbook object, and obtain the worksheet at index 0 with Aspose.Cells in C#. | Open a .xlsx file via a stream, access its first worksheet, and print the worksheet name using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# load workbook from FileStream example | how to get first worksheet from workbook loaded from stream using Aspose.Cells | read Excel file with FileStream and access worksheet index 0 in .NET | C# Aspose.Cells retrieve worksheet name after opening file via stream
// Tags: load workbook from filestream Aspose.Cells | access first worksheet index 0 Aspose.Cells | read excel file via stream C# | retrieve worksheet name Aspose.Cells | open xlsx using FileStream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program opens "sample.xlsx" with a FileStream, loads it into an Aspose.Cells Workbook, accesses the first worksheet (index 0), and writes the worksheet's name to the console.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "sample.xlsx";

        // Open a file stream for reading
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            // Load the workbook from the stream
            Workbook workbook = new Workbook(stream);

            // Access the first worksheet (index 0)
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Example usage: output the name of the first worksheet
            Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
        }
    }
}
