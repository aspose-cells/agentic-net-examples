using System;
using Aspose.Cells;

namespace AsposeCellsXlsxSample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            // Add sample data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["C4"].PutValue(0.6);

            // Save the workbook as an XLSX file using the SaveFormat enumeration
            workbook.Save("SampleOutput.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved as SampleOutput.xlsx");
        }
    }
}