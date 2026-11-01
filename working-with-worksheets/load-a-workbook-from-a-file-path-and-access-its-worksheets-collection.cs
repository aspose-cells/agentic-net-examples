// Title: Load an Excel workbook from a file path and list all worksheet names with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a .xlsx file from a given path using Aspose.Cells and prints each worksheet's name to the console. | Show how to retrieve the Worksheets collection from a Workbook object loaded from disk with Aspose.Cells. | Provide a console application example that reads an Excel file path, loads the workbook, and iterates over worksheets to display their names.
// Common Searches: C# Aspose.Cells how to read worksheet names from a specific file location | example code to load Excel workbook from path and enumerate worksheets using Aspose.Cells | list all sheet names in an .xlsx file with Aspose.Cells .NET library | Aspose.Cells load workbook from disk and get worksheet collection | retrieve worksheet names in a console app using Aspose.Cells for .NET
// Tags: load workbook from file path Aspose.Cells | enumerate worksheet names C# | access Worksheets collection .NET | read Excel sheet names console | Aspose.Cells workbook iteration example

using System;
using Aspose.Cells;

// // Loads an Excel workbook from a specified path with Aspose.Cells, accesses its Worksheets collection, and prints each worksheet's name to the console.
class Program
{
    static void Main()
    {
        // Specify the path to the Excel file
        string filePath = @"C:\Data\SampleWorkbook.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Access the collection of worksheets in the workbook
        WorksheetCollection worksheets = workbook.Worksheets;

        // Example: iterate through the worksheets and print their names
        foreach (Worksheet sheet in worksheets)
        {
            Console.WriteLine($"Worksheet name: {sheet.Name}");
        }
    }
}
