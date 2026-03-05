using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header row
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");

        // Add sample data
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // Save the workbook to an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);

        Console.WriteLine("Workbook saved successfully as output.xlsx");
    }
}