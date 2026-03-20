using System;
using Aspose.Cells;

namespace PrintAreaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue("Data1");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Data2");
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["C1"].PutValue("Extra");
            worksheet.Cells["C2"].PutValue(300);

            // Define the print area (only A1:B3 will be printed)
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Save the workbook with the full content
            workbook.Save("PrintAreaDemo.xlsx");

            // Export only the defined print area to HTML
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true
            };
            workbook.Save("PrintAreaDemo.html", htmlOptions);
        }
    }
}