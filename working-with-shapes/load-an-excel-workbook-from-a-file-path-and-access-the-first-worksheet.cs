using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Specify the path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the given file path using the Workbook(string) constructor
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the workbook (index 0)
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example usage: output the name of the first worksheet
        Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
    }
}