// Title: Generate an HTML <img> tag for a PNG snapshot of an Excel worksheet using Aspose.Cells for .NET (C#)
// Description: The sample creates a workbook, fills cells, configures ImageOrPrintOptions for PNG, renders the first worksheet to a PNG file with SheetRender, and outputs an HTML <img> tag that references the saved image.
// Keywords: Aspose.Cells | C# export Excel to PNG | SheetRender PNG | ImageOrPrintOptions | HTML img tag from worksheet | Excel image embedding | web page Excel snapshot
// Common Searches: Aspose.Cells convert worksheet to PNG C# | how to create HTML img tag for Excel image using Aspose | render Excel sheet as PNG and embed in web page | C# generate PNG from workbook and display in HTML | export Excel to image for web display Aspose.Cells
// Developer Intent: Export a worksheet as a PNG file and produce ready‑to‑use HTML markup that displays the image.
// Use Cases: Show a static view of spreadsheet data on a website without requiring Excel. | Include worksheet snapshots in email newsletters or documentation. | Create printable reports where the layout must stay consistent across browsers.
// AI Prompts: Write C# code with Aspose.Cells that saves the first worksheet as a PNG and prints an <img> tag referencing the file. | Explain the role of ImageOrPrintOptions and SheetRender when converting an Excel sheet to a PNG image. | Show how to customize the generated <img> tag with attributes such as width, height, alt text, and CSS classes.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The sample creates a workbook, fills cells, configures ImageOrPrintOptions for PNG, renders the first worksheet to a PNG file with SheetRender, and outputs an HTML <img> tag that references the saved image.
class Program
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");
        worksheet.Cells["A2"].PutValue(123);
        worksheet.Cells["B2"].PutValue(456);

        // Configure image rendering options for PNG output
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = ImageType.Png;          // PNG format
        imgOptions.OnePagePerSheet = true;             // Render each sheet as a single page

        // Render the worksheet to a PNG file using SheetRender
        SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
        string pngFileName = "worksheet.png";
        sheetRender.ToImage(0, pngFileName);           // Save first (and only) page as PNG

        // Generate an HTML <img> tag that references the saved PNG file
        string htmlImgTag = $"<img src=\"{pngFileName}\" alt=\"Worksheet Image\" />";
        Console.WriteLine(htmlImgTag);
    }
}
