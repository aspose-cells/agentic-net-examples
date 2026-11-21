// Title: Load an Excel workbook from a FileStream and retrieve the first worksheet using Aspose.Cells for .NET
// Common Searches: c# load excel workbook from filestream using aspose.cells | how to get the first worksheet after opening a workbook from a stream in .net | read worksheet name from a workbook loaded via file stream with Aspose.Cells | asp.net console app open xlsx file with FileStream and access worksheet collection | example code for loading .xlsx from a stream and selecting worksheet index 0 Aspose.Cells
// Tags: aspose.cells workbook initialization from stream | aspose.cells first worksheet retrieval | aspose.cells worksheet name extraction | aspose.cells open xlsx stream .net | aspose.cells worksheet collection index access

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example opens an Excel file (sample.xlsx) using a FileStream, loads it into an Aspose.Cells Workbook, accesses the first worksheet (index 0), and prints the worksheet's name to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = "sample.xlsx";

            // Open a file stream for reading the workbook
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook from the stream
                Workbook workbook = new Workbook(fileStream);

                // Access the first worksheet (index 0)
                Worksheet firstWorksheet = workbook.Worksheets[0];

                // Example: output the name of the first worksheet
                Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
            }
        }
    }
}
