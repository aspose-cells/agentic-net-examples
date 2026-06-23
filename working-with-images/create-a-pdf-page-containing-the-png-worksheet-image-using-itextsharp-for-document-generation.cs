using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a workbook and add sample data ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(85);

                // ---------- Save the workbook directly as PDF ----------
                string pdfPath = "WorksheetImage.pdf";
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"PDF created successfully: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}