using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data to cells
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");

        // Save the workbook in XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);

        // Write confirmation to the console
        Console.WriteLine("Workbook saved successfully as 'output.xlsx'.");
    }
}