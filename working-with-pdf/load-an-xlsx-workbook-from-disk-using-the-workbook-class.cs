// Title: Load an XLSX workbook from a local file using Aspose.Cells Workbook constructor in C#
// Common Searches: c# open xlsx workbook from local path using aspose.cells | example code to load excel file with aspose.cells workbook class | how to get first sheet name after loading workbook in aspose.cells c# | asp.net core load spreadsheet from file using aspose.cells | load xlsx file into workbook object with aspose.cells c#
// Tags: Aspose.Cells workbook constructor file path | C# open XLSX workbook using Aspose.Cells | retrieve first worksheet name Aspose.Cells | load Excel file from disk with Aspose.Cells | Aspose.Cells open spreadsheet from local path

using System;
using Aspose.Cells;

// The example demonstrates creating a Workbook instance by passing a full file path to an .xlsx file, which loads the workbook from disk, and then prints the name of the first worksheet to the console.
class Program
{
    static void Main()
    {
        // Specify the full path to the XLSX file you want to load.
        string filePath = @"C:\Path\To\YourWorkbook.xlsx";

        // Load the workbook from disk using the Workbook class constructor.
        Workbook workbook = new Workbook(filePath);

        // At this point the workbook is loaded and ready for further processing.
        // Example: output the name of the first worksheet.
        Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
    }
}
