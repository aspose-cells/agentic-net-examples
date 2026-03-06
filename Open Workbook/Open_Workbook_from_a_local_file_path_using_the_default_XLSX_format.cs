using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Specify the path to the existing Excel file (XLSX is the default format)
        string filePath = "input.xlsx";

        // Open the workbook using the string constructor (default XLSX format)
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Display worksheet name and the value of cell A1
        Console.WriteLine("First worksheet name: " + sheet.Name);
        Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

        // Save the workbook to a new file (demonstrates the Save rule)
        workbook.Save("output.xlsx");
    }
}