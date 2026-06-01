using System;
using Aspose.Cells;

class LoadWorkbookExample
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the specified file using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet (zero‑based index)
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example usage: output the name of the first worksheet
        Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
    }
}