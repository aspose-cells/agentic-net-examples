using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file (default format is assumed to be XLSX)
        string filePath = "input.xlsx";

        // Load the workbook using the string constructor (default format Xlsx)
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Output worksheet name and value of cell A1
        Console.WriteLine("Worksheet Name: " + worksheet.Name);
        Console.WriteLine("Cell A1 Value: " + worksheet.Cells["A1"].StringValue);
    }
}