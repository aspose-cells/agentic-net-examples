// Title: Load an .xlsx file into an Aspose.Cells Workbook in C# and display the first worksheet name and cell A1 value
// AI Prompts: Write C# code that uses the Aspose.Cells Workbook(string) constructor to open a specified .xlsx file and prints the name of the first worksheet together with the text in cell A1. | Show how to wrap the workbook loading and cell reading in a try‑catch block to handle file‑not‑found and format exceptions when using Aspose.Cells. | Adapt the example to load the Excel file from a MemoryStream instead of a file path, then output the same worksheet name and cell value.
// Common Searches: asp.net core open .xlsx using Aspose.Cells and get first sheet name | c# read value of cell A1 after loading Excel file with Aspose.Cells | sample code for Aspose.Cells Workbook(string) constructor | how to retrieve worksheet name and cell content from a loaded workbook in C# | load Excel file into memory with Aspose.Cells without writing to disk
// Tags: file path workbook loading Aspose.Cells C# | first worksheet name retrieval Aspose.Cells | cell A1 extraction Aspose.Cells | Workbook constructor with file path Aspose.Cells | memory stream workbook loading Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates loading an .xlsx file into an Aspose.Cells Workbook, accessing the first worksheet, and printing its name along with the value of cell A1.
class LoadWorkbookDemo
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the specified file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(filePath);

        // Example: access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: read and display the worksheet name and the value of cell A1
        Console.WriteLine("Worksheet Name: " + worksheet.Name);
        Console.WriteLine("Cell A1 Value: " + worksheet.Cells["A1"].StringValue);
    }
}
