using System;
using Aspose.Cells;

namespace AsposeCellsXlsxDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Save the workbook in XLSX format
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Output confirmation to the console
            Console.WriteLine($"Workbook saved successfully as '{outputPath}'.");
        }
    }
}