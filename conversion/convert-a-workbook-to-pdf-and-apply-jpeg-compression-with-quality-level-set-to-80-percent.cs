// Title: C# – Convert an Excel Workbook to PDF with 80% JPEG Compression using Aspose.Cells
// Description: Creates or loads a workbook, configures PdfSaveOptions with SetImageResample(96, 80) to apply 96 PPI JPEG compression at 80% quality, and saves the result as a PDF file.
// Keywords: Aspose.Cells PDF conversion | C# Excel to PDF | JPEG compression 80% quality | SetImageResample | PdfSaveOptions image resampling | reduce PDF file size | export Excel as PDF
// Common Searches: Aspose.Cells export Excel to PDF with JPEG quality | C# SetImageResample PDF image compression | How to lower PDF size when converting Excel with Aspose | Save workbook as PDF with 80% JPEG quality C#
// Developer Intent: Export an Excel workbook to PDF while compressing embedded images to 80% JPEG quality.
// Use Cases: Email‑friendly PDF reports from large spreadsheets. | Generating lightweight invoices or receipts for web portals. | Batch converting multiple workbooks with a uniform compression setting to meet storage limits.
// AI Prompts: Write C# code that loads an existing .xlsx file and saves it as a PDF using Aspose.Cells with JPEG compression set to 80% quality and 96 PPI. | Explain the impact of the SetImageResample parameters on PDF image resolution and file size, and suggest alternatives for higher or lower quality. | Show how to apply different JPEG quality levels per worksheet when exporting a single workbook to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    // Creates or loads a workbook, configures PdfSaveOptions with SetImageResample(96, 80) to apply 96 PPI JPEG compression at 80% quality, and saves the result as a PDF file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B2"].PutValue(123.45);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set JPEG image resampling with desired PPI and quality (80%)
            // Here we use 96 PPI (email quality) as an example
            pdfOptions.SetImageResample(96, 80);

            // Save the workbook as PDF using the specified options
            workbook.Save("ConvertedWorkbook.pdf", pdfOptions);
        }
    }
}
