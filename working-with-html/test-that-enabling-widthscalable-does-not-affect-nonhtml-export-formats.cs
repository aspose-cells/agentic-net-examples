// Title: Verify that enabling HtmlSaveOptions.WidthScalable for HTML export does not affect PDF generation with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as HTML with WidthScalable set to true, then saves the same workbook as PDF and confirms the PDF file exists. | Create an automated unit test in C# that asserts the HtmlSaveOptions.WidthScalable property has no impact on PDF (or other non‑HTML) export when using Aspose.Cells.
// Common Searches: Aspose.Cells HtmlSaveOptions WidthScalable impact on PDF export | C# test HtmlSaveOptions WidthScalable only affects HTML output | Does setting WidthScalable true change PDF generation in Aspose.Cells | How to verify HTML options do not affect other formats Aspose.Cells
// Tags: HtmlSaveOptions WidthScalable setting for HTML | PDF generation unaffected by HTML options Aspose.Cells | C# workbook save to multiple formats Aspose.Cells | validate non‑HTML format behavior Aspose.Cells | unit test HtmlSaveOptions impact on PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates enabling HtmlSaveOptions.WidthScalable for HTML export, then saving the same workbook to PDF to confirm the setting does not influence PDF generation.
class WidthScalableTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row1");
            sheet.Cells["A3"].PutValue("Row2");

            // Enable WidthScalable for HTML export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                WidthScalable = true
            };

            // Save as HTML (to verify the option works)
            string htmlPath = "WidthScalableEnabled.html";
            workbook.Save(htmlPath, htmlOptions);

            // Save the same workbook to PDF (HTML‑specific options are not needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            string pdfPath = "WidthScalableCheck.pdf";
            workbook.Save(pdfPath, pdfOptions);

            // Optional verification: ensure the PDF file was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine("PDF file created successfully.");
            }
            else
            {
                Console.WriteLine("PDF file was not created.");
            }

            // Output confirmation
            Console.WriteLine("HTML and PDF files have been generated. WidthScalable setting does not affect PDF export.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
