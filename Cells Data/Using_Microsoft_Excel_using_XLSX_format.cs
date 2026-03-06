using System;
using Aspose.Cells;

namespace AsposeCellsXlsxDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to cells
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Save the workbook as an XLSX file
            workbook.Save("SampleReport.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved as SampleReport.xlsx");
        }
    }
}