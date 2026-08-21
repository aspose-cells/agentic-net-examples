// Title: C# – Preserve Excel Cell Borders and Gridlines When Converting to PDF with Aspose.Cells
// Description: Demonstrates how to create a workbook, apply a thick blue border style to a range, enable worksheet gridlines, set PageSetup.PrintGridlines, and save the sheet as a PDF using PdfSaveOptions so that both custom borders and gridlines appear exactly as they do in Excel.
// Keywords: Aspose.Cells PDF export | preserve borders PDF | print gridlines Aspose.Cells | C# Excel to PDF | PdfSaveOptions gridlines | cell border style PDF | Aspose.Cells .NET example
// Common Searches: keep cell borders when exporting Excel to PDF Aspose.Cells | enable gridlines in PDF output using Aspose.Cells C# | apply custom border style to range and save as PDF | Aspose.Cells print gridlines PDFSaveOptions | C# convert worksheet to PDF with borders and gridlines
// Developer Intent: Generate a PDF that visually matches the worksheet, including all custom borders and visible gridlines.
// Use Cases: Produce printable reports where table borders and gridlines must remain intact. | Create invoices or statements that retain the exact Excel layout in PDF form. | Archive spreadsheets as PDFs with full visual fidelity for compliance or record‑keeping.
// AI Prompts: Show how to customize PdfSaveOptions (margins, image quality, compression) while preserving borders and gridlines. | Provide code to merge multiple worksheets into a single PDF, keeping each sheet’s borders and gridlines. | Explain how to programmatically toggle gridline visibility and modify border styles before PDF conversion with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a workbook, apply a thick blue border style to a range, enable worksheet gridlines, set PageSetup.PrintGridlines, and save the sheet as a PDF using PdfSaveOptions so that both custom borders and gridlines appear exactly as they do in Excel.
class PreserveBordersAndGridlinesPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item 1");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Item 2");
            worksheet.Cells["B3"].PutValue(200);

            // Define a thick blue border style
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.TopBorder].Color = Color.Blue;
            borderStyle.Borders[BorderType.BottomBorder].Color = Color.Blue;
            borderStyle.Borders[BorderType.LeftBorder].Color = Color.Blue;
            borderStyle.Borders[BorderType.RightBorder].Color = Color.Blue;

            // Apply the border style to the range A1:B3
            AsposeRange range = worksheet.Cells.CreateRange("A1:B3");
            range.ApplyStyle(borderStyle, new StyleFlag { All = true });

            // Make gridlines visible in the worksheet UI and ensure they are printed
            worksheet.IsGridlinesVisible = true;
            worksheet.PageSetup.PrintGridlines = true;

            // Configure PDF save options (gridlines will be printed because of PageSetup)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output file path
            string outputPath = "PreservedBordersGridlines.pdf";

            // Delete existing file if it exists to avoid exceptions
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
