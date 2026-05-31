using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the file path using the Workbook(string) constructor
        Workbook workbook = new Workbook(filePath);

        // The workbook is now loaded and ready for pivot table operations
        Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");
    }
}