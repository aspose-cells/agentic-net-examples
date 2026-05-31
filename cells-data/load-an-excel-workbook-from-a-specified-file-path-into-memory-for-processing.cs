using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Specify the path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the file into memory
        Workbook workbook = new Workbook(filePath);

        // Example processing: access the first worksheet and display its name
        Worksheet firstSheet = workbook.Worksheets[0];
        Console.WriteLine($"First worksheet name: {firstSheet.Name}");

        // Additional processing can be performed here using the 'workbook' object
    }
}