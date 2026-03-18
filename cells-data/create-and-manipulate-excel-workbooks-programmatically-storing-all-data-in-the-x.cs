using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook instance (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the default first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to cells
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(2.50);

            // Save the workbook to disk (lifecycle: save)
            string outputPath = "SampleReport.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}