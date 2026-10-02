// Title: Load an Excel workbook from a file and retrieve the first worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a Workbook instance from a local .xlsx file and accesses the worksheet at index 0 with Aspose.Cells. | Show how to read and print the name of the first worksheet after opening an Excel file using Aspose.Cells in a .NET console app. | Demonstrate initializing a Workbook from a file path and obtaining the first Worksheet object for further processing in C#.
// Common Searches: aspnet aspose.cells how to open an existing .xlsx file and get the first sheet | c# load workbook from disk and read first worksheet name using Aspose.Cells | example code for accessing worksheet index 0 after loading Excel workbook with Aspose.Cells .NET
// Tags: Aspose.Cells load .xlsx workbook C# | access first worksheet by index Aspose.Cells | retrieve worksheet name using Aspose.Cells .NET | create Workbook object from file path Aspose.Cells | console application example Aspose.Cells loading workbook

using Aspose.Cells;
using System;

// // Loads 'input.xlsx' with Aspose.Cells, accesses the first worksheet, and prints its name.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example usage: output the name of the first worksheet
        Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
    }
}
