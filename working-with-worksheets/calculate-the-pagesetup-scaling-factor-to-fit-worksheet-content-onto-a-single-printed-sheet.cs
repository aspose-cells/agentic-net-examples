// Title: C# – Retrieve the scaling factor for fitting a worksheet to a single printed page using Aspose.Cells
// Description: This example creates a workbook, fills it with data, applies PageSetup.SetFitToPages(1,1) and disables percent scaling, then uses SheetRender to read the automatically calculated PageScale (0‑1) and prints it as a percentage. The workbook can be saved after the calculation.
// Keywords: Aspose.Cells C# fit to one page | SheetRender PageScale | SetFitToPages scaling factor | calculate print scale Aspose.Cells | page setup scaling .NET
// Common Searches: Aspose.Cells get scaling percentage after SetFitToPages | how to read page scale for fit‑to‑page printing in .NET | SheetRender.PageScale example | fit entire worksheet on one printed sheet Aspose.Cells
// Developer Intent: Find out the exact print‑scale value that Aspose.Cells applies when a worksheet is configured to fit on one page.
// Use Cases: Display the calculated shrinkage to users before printing. | Log the scale factor for debugging layout issues in automated report pipelines. | Adjust image or PDF export dimensions based on the retrieved PageScale.
// AI Prompts: Show C# code that sets SetFitToPages(1,1), disables percent scaling, and outputs SheetRender.PageScale as a percentage. | Explain the algorithm behind SheetRender.PageScale and how to use it to modify rendering options before exporting. | Generate a snippet that saves the workbook after fitting it to one page and writes the scaling factor to a log file.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates a workbook, fills it with data, applies PageSetup.SetFitToPages(1,1) and disables percent scaling, then uses SheetRender to read the automatically calculated PageScale (0‑1) and prints it as a percentage. The workbook can be saved after the calculation.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill the sheet with enough data to normally span multiple pages
        for (int row = 0; row < 100; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Configure page setup to fit the entire sheet onto a single printed page
        PageSetup pageSetup = sheet.PageSetup;
        // Use the SetFitToPages method (rule) to specify 1 page wide and 1 page tall
        pageSetup.SetFitToPages(1, 1);
        // Ensure scaling is driven by FitToPages rather than a percent zoom
        pageSetup.IsPercentScale = false;

        // Create rendering options (default settings)
        ImageOrPrintOptions options = new ImageOrPrintOptions();

        // Create SheetRender after page‑setup changes (rule)
        SheetRender sheetRender = new SheetRender(sheet, options);

        // Retrieve the calculated page scale (0.0‑1.0 range)
        double pageScale = sheetRender.PageScale;

        // Output the scaling factor as a percentage
        Console.WriteLine($"Calculated page scale to fit on one sheet: {pageScale * 100:0.##}%");

        // Save the workbook (optional, demonstrates lifecycle rule)
        workbook.Save("FitToOnePage.xlsx");
    }
}
