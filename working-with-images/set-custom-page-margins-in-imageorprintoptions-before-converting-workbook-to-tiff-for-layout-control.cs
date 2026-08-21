// Title: Set Custom Page Margins with ImageOrPrintOptions and Export Worksheet to TIFF (C# Aspose.Cells)
// Description: Demonstrates how to define left, right, top, and bottom margins in centimeters via Worksheet.PageSetup, configure ImageOrPrintOptions for TIFF output, enable OnePagePerSheet, and render the sheet to a single‑page TIFF file using SheetRender.
// Keywords: Aspose.Cells C# | ImageOrPrintOptions TIFF | custom page margins | OnePagePerSheet | SheetRender example | Excel to TIFF conversion | margin settings centimeters
// Common Searches: Aspose.Cells set page margins before TIFF export | C# render Excel sheet as single page TIFF | ImageOrPrintOptions custom margins example | how to use OnePagePerSheet with TIFF | convert workbook to TIFF with specific margins
// Developer Intent: Apply precise margin dimensions to a worksheet and generate a single‑page TIFF image using Aspose.Cells for .NET.
// Use Cases: Produce printable reports where margin layout must match corporate standards. | Archive Excel data as high‑resolution TIFF files with exact page formatting. | Create thumbnail previews of spreadsheets for document management systems.
// AI Prompts: Generate C# code that sets left, right, top, and bottom margins in centimeters and saves the worksheet as a one‑page TIFF using Aspose.Cells. | Explain the impact of ImageOrPrintOptions.OnePagePerSheet on TIFF output when custom margins are defined. | Provide a step‑by‑step tutorial for adjusting page margins and exporting a worksheet to a TIFF file with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to define left, right, top, and bottom margins in centimeters via Worksheet.PageSetup, configure ImageOrPrintOptions for TIFF output, enable OnePagePerSheet, and render the sheet to a single‑page TIFF file using SheetRender.
class SetCustomMarginsAndRenderTiff
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample content
        sheet.Cells["A1"].PutValue("Custom Margin Demo");
        sheet.Cells["A2"].PutValue("This page uses custom margins.");

        // Set custom page margins (centimeters)
        sheet.PageSetup.LeftMargin = 2.0;    // left margin
        sheet.PageSetup.RightMargin = 2.0;   // right margin
        sheet.PageSetup.TopMargin = 1.5;     // top margin
        sheet.PageSetup.BottomMargin = 1.5;  // bottom margin

        // Configure image options for TIFF rendering
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = Aspose.Cells.Drawing.ImageType.Tiff;
        options.OnePagePerSheet = true; // render the whole sheet on a single page

        // Create a SheetRender with the worksheet and options
        SheetRender renderer = new SheetRender(sheet, options);

        // Render the worksheet to a TIFF file
        using (FileStream tiffStream = new FileStream("CustomMarginsOutput.tiff", FileMode.Create))
        {
            renderer.ToTiff(tiffStream);
        }

        Console.WriteLine("TIFF file generated with custom margins.");
    }
}
