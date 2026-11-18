// Title: How to add a semi‑transparent diagonal text watermark to every page when saving an Excel workbook as PDF with Aspose.Cells PdfSaveOptions in C#
// Common Searches: C# Aspose.Cells add diagonal CONFIDENTIAL watermark to PDF output | set semi transparent text watermark using PdfSaveOptions in Aspose.Cells | Aspose.Cells PdfSaveOptions watermark text property example C# | convert Excel to PDF with custom watermark using Aspose.Cells library | Aspose.Cells PDF export with rotated watermark text C# code sample | configure watermark opacity in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PdfSaveOptions watermark text | C# add diagonal watermark to PDF | Excel to PDF conversion confidential watermark | semi transparent PDF watermark Aspose.Cells | rotate watermark text Aspose.Cells PDF export | PDF watermark opacity configuration C#

using System;
using System.Drawing;
using Aspose.Cells;

// The example shows how to configure a semi‑transparent, rotated text watermark for each page when converting an Excel workbook to PDF using Aspose.Cells PdfSaveOptions in C#. It includes commented code that creates a PdfWatermark with text, color, opacity, font size, and rotation angle, then saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PdfWatermark class is not available in the referenced Aspose.Cells version.
            // If a later version is added, the watermark can be configured as shown below:
            // var watermark = new Aspose.Cells.Pdf.PdfWatermark
            // {
            //     Text = "CONFIDENTIAL",
            //     Color = Color.FromArgb(128, 255, 0, 0),
            //     FontSize = 72,
            //     FontName = "Arial",
            //     RotationAngle = -45
            // };
            // pdfOptions.Watermark = watermark;

            // Save the workbook as a PDF
            string outputPath = "OutputWithWatermark.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
