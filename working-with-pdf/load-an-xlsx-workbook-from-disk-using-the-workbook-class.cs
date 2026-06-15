using System;
using Aspose.Cells;

class LoadWorkbookDemo
{
    static void Main()
    {
        // Path to the Excel file on disk
        string filePath = "example.xlsx";

        // Load the workbook using the string constructor (Workbook(string))
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Output worksheet information
        Console.WriteLine($"Worksheet Name: {worksheet.Name}");
        Console.WriteLine($"Number of Cells: {worksheet.Cells.Count}");

        // Example: display the value of cell A1 if it exists
        var cell = worksheet.Cells["A1"];
        if (cell.Value != null)
        {
            Console.WriteLine($"A1 Value: {cell.Value}");
        }
    }
}