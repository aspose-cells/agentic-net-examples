using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the XLSX workbook (Excel 2007‑2019 format)
        string workbookPath = "SampleWorkbook.xlsx";

        // Load the workbook using the constructor that accepts a file path
        Workbook workbook = new Workbook(workbookPath);

        // Example operation: display the number of worksheets in the loaded workbook
        Console.WriteLine($"Workbook '{workbookPath}' loaded successfully.");
        Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

        // Optional: access the first worksheet and read a cell value
        Worksheet firstSheet = workbook.Worksheets[0];
        Console.WriteLine($"First worksheet name: {firstSheet.Name}");
        Console.WriteLine($"Value of cell A1: {firstSheet.Cells["A1"].StringValue}");
    }
}