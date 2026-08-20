// Title: Render Excel Chart to High‑Resolution PNG and Save as PDF with a Custom Color Profile using Aspose.Cells for .NET
// Description: This example creates a workbook with a column chart, renders the chart to a 300 DPI PNG, inserts the PNG into a new workbook, and saves the workbook as a PDF. The PDF is generated with PdfSaveOptions.ImageResample set to 300 PPI and JPEG quality 100, effectively preserving the chart's color fidelity for accurate printing.
// Keywords: Aspose.Cells | C# | Excel chart to PDF | high DPI PNG | custom color profile | PdfSaveOptions image resample | printing color fidelity | ImageOrPrintOptions | chart image embedding | ICC profile simulation
// Common Searches: Aspose.Cells export chart to PDF with custom color profile | how to set image resample DPI in PdfSaveOptions C# | render Excel chart as high‑resolution PNG for printing | preserve chart colors when converting Excel to PDF | C# Aspose.Cells embed PNG in PDF with 300 DPI
// Developer Intent: I need to convert an Excel chart to a PDF while maintaining exact colors for print, using a high‑resolution image or custom color profile via Aspose.Cells.
// Use Cases: Producing print‑ready financial reports where chart colors must match corporate branding. | Generating marketing brochures from Excel data with precise color reproduction for pre‑press. | Automating batch conversion of Excel dashboards to PDFs that retain color accuracy for high‑quality printing.
// AI Prompts: Show C# code that applies an ICC color profile to chart images when saving a workbook as PDF with Aspose.Cells. | Explain how PdfSaveOptions.ImageResample and JPEG quality settings simulate a custom color profile for printed PDFs. | Provide a sample that embeds multiple high‑resolution chart images, each with individual resampling settings, into a single PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomColorProfileDemo
{
    // This example creates a workbook with a column chart, renders the chart to a 300 DPI PNG, inserts the PNG into a new workbook, and saves the workbook as a PDF. The PDF is generated with PdfSaveOptions.ImageResample set to 300 PPI and JPEG quality 100, effectively preserving the chart's color fidelity for accurate printing.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook and add sample data for a chart
                // ------------------------------------------------------------
                Workbook sourceWb = new Workbook();
                Worksheet srcWs = sourceWb.Worksheets[0];

                srcWs.Cells["A1"].PutValue("Category");
                srcWs.Cells["A2"].PutValue("Apple");
                srcWs.Cells["A3"].PutValue("Orange");
                srcWs.Cells["A4"].PutValue("Banana");

                srcWs.Cells["B1"].PutValue("Value");
                srcWs.Cells["B2"].PutValue(120);
                srcWs.Cells["B3"].PutValue(80);
                srcWs.Cells["B4"].PutValue(150);

                // Add a column chart
                int chartIdx = srcWs.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = srcWs.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // 2. Render the chart to a high‑resolution PNG image
                // ------------------------------------------------------------
                string chartImagePath = "ChartImage.png";

                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,          // PNG preserves color fidelity
                    HorizontalResolution = 300,         // 300 DPI for printing quality
                    VerticalResolution = 300
                };

                // Render chart to image file
                chart.ToImage(chartImagePath, imgOptions);

                // ------------------------------------------------------------
                // 3. Create a new workbook and embed the rendered chart image
                // ------------------------------------------------------------
                Workbook targetWb = new Workbook();
                Worksheet tgtWs = targetWb.Worksheets[0];

                // Ensure the image file exists before inserting
                if (!File.Exists(chartImagePath))
                    throw new FileNotFoundException("Rendered chart image not found.", chartImagePath);

                // Insert the image as a picture shape
                int picIdx = tgtWs.Pictures.Add(0, 0, chartImagePath);
                Picture pic = tgtWs.Pictures[picIdx];

                // ------------------------------------------------------------
                // 4. Save the workbook as PDF with image resampling (simulating a color profile)
                // ------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Resample images to 300 PPI and use high JPEG quality (100%)
                pdfOptions.SetImageResample(300, 100);

                string outputPdf = "ChartWithCustomColorProfile.pdf";
                targetWb.Save(outputPdf, pdfOptions);

                Console.WriteLine("PDF generated successfully: " + Path.GetFullPath(outputPdf));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
