// Title: Load an Excel workbook from a file and retrieve the first worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to open a .xlsx file from a specified path and obtain the worksheet at index 0. | Create a console application that loads a workbook from disk with Aspose.Cells, accesses the first sheet, and prints its name.
// Common Searches: Aspose.Cells C# open .xlsx file and get first sheet name | How to read the first worksheet of an Excel workbook using Aspose.Cells in .NET | C# example for loading workbook from file path with Aspose.Cells and accessing worksheet index 0 | Retrieve worksheet name from Excel file using Aspose.Cells library
// Tags: load workbook from file Aspose.Cells C# | access first worksheet by index Aspose.Cells | read worksheet name .xlsx C# | Aspose.Cells open workbook by path

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Loads "input.xlsx" using Aspose.Cells, accesses the first worksheet (index 0), and writes its name to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (index 0)
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Example: output the name of the first worksheet
            Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
        }
    }
}
