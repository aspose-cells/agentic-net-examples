// Title: Saving an Aspose.Cells workbook to PDF with 80% JPEG image quality in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to export a workbook to PDF while setting the image resample PPI to 96 and JPEG quality to 80%. | Show how to configure PdfSaveOptions in Aspose.Cells for .NET to balance PDF file size and clarity by adjusting image quality. | Provide a step‑by‑step example of calling SetImageResample on PdfSaveOptions to control image compression during workbook‑to‑PDF conversion.
// Common Searches: Aspose.Cells C# export workbook to PDF with specific JPEG quality | Set image resampling options when saving PDF with Aspose.Cells .NET | How to reduce PDF size from Aspose.Cells by adjusting image quality
// Tags: Aspose.Cells PDF export with image quality settings | C# PdfSaveOptions SetImageResample usage | balance PDF file size and clarity Aspose.Cells | configure JPEG compression Aspose.Cells PDF | workbook to PDF conversion image optimization

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    // The example creates a workbook, adds sample data, configures PdfSaveOptions to resample images at 96 PPI with 80 % JPEG quality, and saves the workbook as output.pdf.
    public class WorkbookToPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Text");
                sheet.Cells["B2"].PutValue(123.45);
                sheet.Cells["C3"].PutValue(DateTime.Now);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Set image resampling: desired PPI (e.g., 96) and JPEG quality to 80%
                pdfOptions.SetImageResample(96, 80); // 80% quality for balanced size and clarity

                // Save the workbook as PDF using the configured options
                workbook.Save("output.pdf", pdfOptions);
                Console.WriteLine("PDF file 'output.pdf' has been created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToPdf.Run();
        }
    }
}
