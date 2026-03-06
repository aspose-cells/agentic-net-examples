using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfAdvancedDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruits");
            sheet.Cells["A3"].PutValue("Vegetables");
            sheet.Cells["A4"].PutValue("Grains");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // Add a column chart based on the data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Data Chart";

            // -------------------- PDF Save Options --------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set PDF/A-2b compliance (archival)
            pdfOptions.Compliance = PdfCompliance.PdfA2b;

            // Optimize for minimum file size
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Use Flate compression for streams
            pdfOptions.PdfCompression = PdfCompressionCore.Flate;

            // Embed standard Windows fonts (helps with PDF/A compliance)
            pdfOptions.EmbedStandardWindowsFonts = true;

            // Set a custom default font for Unicode characters
            pdfOptions.DefaultFont = "Arial Unicode MS";

            // Enable checking of font compatibility (helps catch missing glyphs)
            pdfOptions.CheckFontCompatibility = true;

            // Add a watermark to the PDF
            RenderingFont watermarkFont = new RenderingFont("Calibri", 72)
            {
                Bold = true,
                Italic = true,
                Color = Color.FromArgb(128, Color.Gray) // semi‑transparent gray
            };
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                Rotation = 45f,
                Opacity = 0.3f,
                ScaleToPagePercent = 80
            };
            pdfOptions.Watermark = watermark;

            // Configure security options: password protection and permissions
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                OwnerPassword = "ownerPwd123",
                UserPassword = "userPwd123",
                PrintPermission = true,               // allow printing
                FullQualityPrintPermission = true,    // allow high‑quality printing
                FillFormsPermission = true            // allow filling interactive forms
            };
            pdfOptions.SecurityOptions = security;

            // Save the workbook as a PDF with the configured options
            workbook.Save("AdvancedOutput.pdf", pdfOptions);

            // -------------------- Export Chart Directly to PDF --------------------
            // The chart can also be saved as a separate PDF file
            chart.ToPdf("ChartOnly.pdf");

            // -------------------- Print to Physical Printer (optional) --------------------
            // Demonstrates rendering to a printer; replace with an actual printer name if needed
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                // Print only the first page of the PDF rendering
                PageIndex = 0,
                PageCount = 1
            };
            WorkbookRender renderer = new WorkbookRender(workbook, printOptions);
            // Uncomment the line below and provide a valid printer name to actually print
            // renderer.ToPrinter("Microsoft Print to PDF");

            Console.WriteLine("PDF generation completed with advanced settings.");
        }
    }
}