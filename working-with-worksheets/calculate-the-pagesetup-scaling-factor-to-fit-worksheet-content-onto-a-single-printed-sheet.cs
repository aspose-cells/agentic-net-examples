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

        // Populate some sample data (optional, just to have content)
        for (int i = 0; i < 30; i++)
        {
            sheet.Cells[i, 0].PutValue($"Item {i + 1}");
            sheet.Cells[i, 1].PutValue((i + 1) * 10);
        }

        // Configure page setup to fit everything on a single printed page
        PageSetup pageSetup = sheet.PageSetup;
        pageSetup.IsPercentScale = false;          // Use FitToPages rather than a fixed Zoom
        pageSetup.SetFitToPages(1, 1);              // Fit to 1 page wide and 1 page tall
        pageSetup.PrintArea = $"A1:B{sheet.Cells.MaxDataRow + 1}"; // Define the print area

        // Create rendering options (default options are sufficient)
        ImageOrPrintOptions options = new ImageOrPrintOptions();

        // Create SheetRender after page‑setup changes
        SheetRender sheetRender = new SheetRender(sheet, options);

        // Retrieve the calculated page scale (e.g., 0.75 = 75%)
        double pageScale = sheetRender.PageScale;

        // Output the scaling factor as a percentage
        Console.WriteLine($"Calculated page scale to fit on one sheet: {pageScale * 100}%");

        // Save the workbook (demonstrates use of the save rule)
        workbook.Save("FitOnePage.xlsx");
    }
}