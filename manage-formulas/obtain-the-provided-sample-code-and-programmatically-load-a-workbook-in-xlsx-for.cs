using System;
using Aspose.Cells;

class LoadWorkbookDemo
{
    public static void Main()
    {
        // Path to the XLSX file to be loaded
        string filePath = "sample.xlsx";

        // Load the workbook using the string constructor (Workbook(string))
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Output basic information about the loaded workbook
        Console.WriteLine("Worksheet Name: " + worksheet.Name);
        Console.WriteLine("Cell A1 Value: " + worksheet.Cells["A1"].StringValue);
    }
}