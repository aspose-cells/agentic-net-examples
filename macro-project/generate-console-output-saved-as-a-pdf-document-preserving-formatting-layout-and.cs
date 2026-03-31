using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ----- Populate worksheet with sample data -----
            // Header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["C1"].PutValue("Price");
            worksheet.Cells["D1"].PutValue("Total");

            // Sample data rows
            var items = new[]
            {
                new { Product = "Apple",  Quantity = 10, Price = 0.5 },
                new { Product = "Banana", Quantity = 5,  Price = 0.3 },
                new { Product = "Cherry", Quantity = 20, Price = 0.2 }
            };

            int currentRow = 1; // zero‑based index, row 2 in Excel
            foreach (var item in items)
            {
                worksheet.Cells[currentRow, 0].PutValue(item.Product);
                worksheet.Cells[currentRow, 1].PutValue(item.Quantity);
                worksheet.Cells[currentRow, 2].PutValue(item.Price);
                // Calculate total using a formula
                worksheet.Cells[currentRow, 3].Formula = $"=B{currentRow + 1}*C{currentRow + 1}";
                currentRow++;
            }

            // ----- Apply formatting to header -----
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            StyleFlag styleFlag = new StyleFlag { All = true };
            worksheet.Cells.CreateRange("A1:D1").ApplyStyle(headerStyle, styleFlag);

            // Auto‑fit columns to preserve layout
            worksheet.AutoFitColumns();

            // ----- Configure page setup for PDF output -----
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.Orientation = PageOrientationType.Landscape;
            pageSetup.PaperSize = PaperSizeType.PaperA4;
            pageSetup.FitToPagesWide = 1;   // Fit width to a single page
            pageSetup.FitToPagesTall = 0;   // Let height adjust automatically

            // ----- Save workbook as PDF preserving formatting -----
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false // Keep original pagination
            };

            string outputPdf = "Report.pdf";
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF: {outputPdf}");
        }
    }
}