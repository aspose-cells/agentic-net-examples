using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsChartFreePdfDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (chart‑free)
                Workbook workbook = new Workbook();

                // Populate sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(85);

                // Apply bold style to header row (range A1:B1)
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                AsposeRange headerRange = sheet.Cells.CreateRange("A1:B1");
                headerRange.SetStyle(headerStyle);

                // Configure PDF save options to retain document structure (preserves formatting)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Save the workbook as PDF
                string outputPath = "ChartFreeWorkbook.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook saved to PDF successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}