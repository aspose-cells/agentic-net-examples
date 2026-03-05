using System;
using Aspose.Cells;

namespace AsposeCellsXlsxExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data to cells
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apples");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["A3"].PutValue("Bananas");
            worksheet.Cells["B3"].PutValue(85);

            // Save the workbook as an XLSX file
            workbook.Save("SampleReport.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook created and saved as SampleReport.xlsx");
        }
    }
}