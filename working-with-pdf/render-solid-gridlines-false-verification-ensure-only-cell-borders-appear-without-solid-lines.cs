// Title: C# – Render Excel to PNG with Gridlines Hidden, Showing Only Cell Borders using Aspose.Cells
// Description: Shows how to build a workbook, apply thin borders, turn off the default gridlines, set GridlineType to Hair, and export the first sheet as a PNG image with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# hide gridlines | render Excel to PNG | gridline type hair | cell borders only | export worksheet image | disable default gridlines | ImageOrPrintOptions Aspose.Cells | C# Excel image export | Aspose.Cells border style | Excel to PNG without gridlines
// Common Searches: Aspose.Cells hide worksheet gridlines C# | Export Excel sheet to PNG with only borders | GridlineType Hair Aspose.Cells example | C# render Excel as image without solid lines | Apply thin borders to range Aspose.Cells
// Developer Intent: Create a PNG image of an Excel sheet that displays only the explicitly defined cell borders and no default gridlines.
// Use Cases: Generating clean table images for web dashboards where Excel gridlines would be distracting. | Producing printable report graphics that rely on custom border styling. | Automated testing to confirm that GridlineType settings suppress solid gridlines in rendered output.
// AI Prompts: How can I change the border color while keeping gridlines hidden in the PNG export? | Show me the steps to export the same worksheet to PDF with only cell borders visible. | Provide a verification script that checks the rendered PNG contains no solid gridlines.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Shows how to build a workbook, apply thin borders, turn off the default gridlines, set GridlineType to Hair, and export the first sheet as a PNG image with Aspose.Cells for .NET.
class SolidGridlinesVerification
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(20);

            // Define a style that contains thin borders on all sides
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Apply the border style to the range A1:B3
            StyleFlag flag = new StyleFlag { All = true };
            AsposeRange range = worksheet.Cells.CreateRange("A1:B3");
            range.ApplyStyle(borderStyle, flag);

            // Hide the default worksheet gridlines so only the explicit borders are visible
            worksheet.IsGridlinesVisible = false;

            // Configure rendering options: use a non‑solid gridline type (Hair) to ensure no solid lines appear
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                GridlineType = GridlineType.Hair // non‑solid gridline style
            };

            // Render the first sheet to an image file
            SheetRender sheetRender = new SheetRender(worksheet, renderOptions);
            string imagePath = "BordersOnly.png";
            sheetRender.ToImage(0, imagePath);
            Console.WriteLine($"Image saved to: {Path.GetFullPath(imagePath)}");

            // Save the workbook (optional, for verification)
            string workbookPath = "BordersOnly.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
