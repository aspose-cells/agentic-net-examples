using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the sample XLSX file (adjust the folder as needed)
        string dataDir = "Data";
        string fileName = "sample.xlsx";
        string filePath = System.IO.Path.Combine(dataDir, fileName);

        // Open the workbook using the default XLSX format
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet and display some basic information
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine($"Worksheet Name: {sheet.Name}");
        Console.WriteLine($"Number of cells: {sheet.Cells.Count}");

        // Clean up resources
        workbook.Dispose();
    }
}