using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the existing XLSX file
        string filePath = "input.xlsx";

        // Open the workbook using the constructor that accepts a file path
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Display the name of the first worksheet
        Console.WriteLine("First worksheet name: " + sheet.Name);

        // Save a copy of the workbook (optional)
        workbook.Save("copy.xlsx", SaveFormat.Xlsx);
    }
}