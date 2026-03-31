using System;
using Aspose.Cells;

class LoadWorkbookDemo
{
    static void Main()
    {
        // Path to the source Excel file (can be .xlsx, .xls, .csv, .xml, etc.)
        string filePath = "input.xlsx";

        // Load the workbook using the constructor that accepts a file path
        Workbook workbook = new Workbook(filePath);

        // Display basic information to confirm successful loading
        Console.WriteLine($"Workbook loaded. Worksheets count: {workbook.Worksheets.Count}");
        Console.WriteLine($"First worksheet name: {workbook.Worksheets[0].Name}");

        // The workbook is now ready for subsequent XML map import operations
    }
}