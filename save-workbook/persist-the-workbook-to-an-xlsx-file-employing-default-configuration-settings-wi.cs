using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Optional: add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Save the workbook to an XLSX file using default configuration
        workbook.Save("output.xlsx");

        // Release resources
        workbook.Dispose();
    }
}