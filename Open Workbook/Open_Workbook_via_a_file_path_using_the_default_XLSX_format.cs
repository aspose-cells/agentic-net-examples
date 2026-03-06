using System;
using Aspose.Cells;

namespace AsposeCellsOpenWorkbookDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file (XLSX format is the default)
            string filePath = "input.xlsx";

            // Open the workbook using the file path constructor
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Display basic information about the opened workbook
            Console.WriteLine($"Workbook opened from: {filePath}");
            Console.WriteLine($"First worksheet name: {worksheet.Name}");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
        }
    }
}