using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Author note: simple example to load an XLSX file
        string filePath = "sample.xlsx";

        // Load the workbook from disk (no special load options needed)
        Workbook workbook = new Workbook(filePath);

        // Demonstrate that the workbook is loaded
        Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
    }
}