// Title: Load an Excel workbook from a file and get the name of the first worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a Workbook from a given .xlsx file path and prints the first worksheet's name to the console using Aspose.Cells. | Show how to open an existing Excel file with Aspose.Cells, retrieve the worksheet at index 0, and display its Name property. | Demonstrate loading a workbook from disk and accessing the first sheet's metadata (e.g., Name) in a .NET console application.
// Common Searches: aspnet load excel file with Aspose.Cells and read first sheet name | c# Aspose.Cells get worksheet name from existing workbook | how to open .xlsx using Aspose.Cells and retrieve sheet at index 0 | example of reading first worksheet name in Aspose.Cells C# console app
// Tags: open .xlsx workbook Aspose.Cells C# | access first worksheet by index Aspose.Cells | read worksheet name property Aspose.Cells | load Excel file from disk Aspose.Cells | retrieve first sheet name Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // The example loads 'input.xlsx' into an Aspose.Cells Workbook, accesses the first worksheet (index 0), and writes its Name to the console.
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file
            string filePath = "input.xlsx";

            // Load the workbook from the file using the string constructor
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (index 0)
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Output the name of the first worksheet
            Console.WriteLine($"First worksheet name: {firstWorksheet.Name}");
        }
    }
}
