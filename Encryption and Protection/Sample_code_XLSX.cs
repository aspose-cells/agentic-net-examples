using System;
using Aspose.Cells;

namespace AsposeCellsSample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["C1"].PutValue("Price");

            // Add sample data rows
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["C2"].PutValue(0.5);

            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["C3"].PutValue(0.3);

            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B4"].PutValue(15);
            worksheet.Cells["C4"].PutValue(0.8);

            // Save the workbook as an XLSX file (lifecycle: save)
            workbook.Save("SampleOutput.xlsx", SaveFormat.Xlsx);

            // Optional: inform the user
            Console.WriteLine("Workbook saved successfully as SampleOutput.xlsx");
        }
    }
}