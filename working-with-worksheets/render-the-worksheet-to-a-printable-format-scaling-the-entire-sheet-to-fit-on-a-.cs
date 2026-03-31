using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and obtain the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill the worksheet with sample data
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Configure the page setup to fit the entire sheet on a single printed page
        PageSetup pageSetup = sheet.PageSetup;
        pageSetup.SetFitToPages(1, 1);               // 1 page wide, 1 page tall
        pageSetup.PrintArea = "A1:J20";              // optional: define the print area

        // Create image/print options; OnePagePerSheet ensures the renderer treats the sheet as one page
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            OnePagePerSheet = true,
            ImageType = Aspose.Cells.Drawing.ImageType.Png
        };

        // Instantiate SheetRender after page setup modifications
        SheetRender renderer = new SheetRender(sheet, options);

        // Render the single page to a PNG file
        string outputFile = "Sheet_FitToOnePage.png";
        renderer.ToImage(0, outputFile);

        // Release resources used by the renderer
        renderer.Dispose();

        Console.WriteLine($"Worksheet rendered to '{outputFile}'.");
    }
}