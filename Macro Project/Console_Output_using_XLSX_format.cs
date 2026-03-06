using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to cells
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");

        // Save the workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);

        // Write confirmation to the console
        Console.WriteLine("Workbook saved successfully as output.xlsx");
    }
}