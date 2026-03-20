using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int row = 0; row < 50; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Configure the page setup to fit the entire sheet on a single page
        // This scales the content automatically.
        sheet.PageSetup.SetFitToPages(1, 1); // 1 page wide, 1 page tall

        // Create image/print options.
        // OnePagePerSheet forces the renderer to treat the whole sheet as one page.
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            OnePagePerSheet = true,
            ImageType = Aspose.Cells.Drawing.ImageType.Png
        };

        // Create the SheetRender object using the worksheet and options
        SheetRender renderer = new SheetRender(sheet, options);

        // Render the (only) page to an image file
        renderer.ToImage(0, "SinglePageOutput.png");

        // Release resources
        renderer.Dispose();
    }
}