using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Save the workbook to a file.
        string filePath = "Sample.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to {filePath}");
    }
}