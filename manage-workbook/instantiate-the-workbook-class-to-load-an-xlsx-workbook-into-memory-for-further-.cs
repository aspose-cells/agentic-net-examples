using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the XLSX file to be loaded
        string filePath = "sample.xlsx";

        // Load the workbook into memory using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Example manipulation: access the first worksheet and read a cell value
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Value of A1: " + sheet.Cells["A1"].StringValue);
    }
}