using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Add a new worksheet with a custom name
        Worksheet worksheet = workbook.Worksheets.Add("MySheet");

        // Example: put a value in the first cell
        worksheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

        // Save the workbook to disk
        workbook.Save("CreatedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}