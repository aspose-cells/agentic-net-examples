using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ResponsiveImageGenerator
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // ---------- 1x (standard DPI) ----------
        ImageOrPrintOptions opts1x = new ImageOrPrintOptions();
        opts1x.ImageType = ImageType.Png;               // PNG output
        // Default DPI is 96, suitable for 1x images

        SheetRender render1x = new SheetRender(sheet, opts1x);
        render1x.ToImage(0, "Sheet_1x.png");            // Render first page to PNG
        render1x.Dispose();

        // ---------- 2x (high DPI) ----------
        ImageOrPrintOptions opts2x = new ImageOrPrintOptions();
        opts2x.ImageType = ImageType.Png;
        opts2x.HorizontalResolution = 192;              // Double the standard DPI
        opts2x.VerticalResolution = 192;

        SheetRender render2x = new SheetRender(sheet, opts2x);
        render2x.ToImage(0, "Sheet_2x.png");            // Render first page to high‑resolution PNG
        render2x.Dispose();
    }
}