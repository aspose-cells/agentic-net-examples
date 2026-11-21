// Title: Set custom page margins when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: how to define custom top and bottom margins for PDF export using Aspose.Cells C# | Aspose.Cells C# set left and right page margins before saving workbook as PDF | C# PdfSaveOptions margin configuration for Excel to PDF conversion with Aspose.Cells | adjust worksheet page setup margins for PDF output in Aspose.Cells .NET | example code to change page margins when exporting Excel to PDF using Aspose.Cells
// Tags: Aspose.Cells worksheet page margin settings | PdfSaveOptions custom margin values C# | Excel to PDF margin adjustment Aspose | C# export workbook with specific page margins | page setup margin configuration Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The sample loads an existing Excel file or creates a new workbook, applies custom top/bottom (0.5 inch) and left/right (1 inch) margins to each worksheet via PageSetup, configures PdfSaveOptions, and saves the workbook as a PDF with the defined margins.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Load workbook if the file exists; otherwise create a simple workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample data");
            }

            // Set custom page margins (values are in points; 1 point = 1/72 inch).
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.TopMargin = 36;    // 0.5 inch
                sheet.PageSetup.BottomMargin = 36; // 0.5 inch
                sheet.PageSetup.LeftMargin = 72;   // 1 inch
                sheet.PageSetup.RightMargin = 72;  // 1 inch
            }

            // Configure PDF save options if needed.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the custom margin settings.
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
