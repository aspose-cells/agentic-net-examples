using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class CustomPaperSizeRender
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to demonstrate the rendering area
        sheet.Cells["A1"].PutValue("Custom Paper Size Demo");
        sheet.Cells["A2"].PutValue("Width: 2 inches, Height: 3 inches");

        // Set a custom paper size (2 inches wide, 3 inches high)
        sheet.PageSetup.CustomPaperSize(2.0, 3.0);
        // Ensure the page setup uses the custom size
        sheet.PageSetup.PaperSize = PaperSizeType.Custom;

        // Configure image/print options for rendering
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,   // Output format
            OnePagePerSheet = true       // Render each sheet as a single page
        };

        // Create a SheetRender instance after page setup modifications
        SheetRender render = new SheetRender(sheet, options);

        // Render the first (and only) page to a PNG file
        render.ToImage(0, "CustomPaperSize.png");

        // Release resources used by SheetRender
        render.Dispose();

        // Save the workbook to verify the custom paper size settings
        workbook.Save("CustomPaperSizeDemo.xlsx");
    }
}