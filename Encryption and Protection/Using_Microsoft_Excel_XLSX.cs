using System;
using Aspose.Cells;

namespace AsposeCellsXlsxDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the default worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to cells
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(200);

            // Save the workbook as an XLSX file (lifecycle: save)
            string outputPath = "SampleOutput.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}