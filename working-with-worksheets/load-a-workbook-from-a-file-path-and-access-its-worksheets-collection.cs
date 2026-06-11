using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "example.xlsx";

        // Load the workbook from the specified file path using the Workbook(string) constructor
        Workbook workbook = new Workbook(filePath);

        // Access the collection of worksheets in the workbook
        WorksheetCollection worksheets = workbook.Worksheets;

        // Example: iterate through the worksheets and output their names
        for (int i = 0; i < worksheets.Count; i++)
        {
            Console.WriteLine($"Worksheet {i}: {worksheets[i].Name}");
        }

        // Clean up resources
        workbook.Dispose();
    }
}