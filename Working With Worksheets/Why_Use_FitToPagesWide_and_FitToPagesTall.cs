using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class FitToPagesDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that spans many columns and rows
        // This will make the sheet require multiple printed pages by default
        for (int row = 0; row < 50; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Get the PageSetup object to configure printing options
        PageSetup pageSetup = sheet.PageSetup;

        // ------------------------------------------------------------
        // Why use FitToPagesWide and FitToPagesTall?
        // ------------------------------------------------------------
        // * FitToPagesWide controls how many pages wide the printed sheet will be.
        //   Setting it to 1 forces all columns to be squeezed onto a single page width.
        // * FitToPagesTall controls how many pages tall the printed sheet will be.
        //   Setting it to 0 tells Excel/Aspose.Cells to calculate the required number
        //   of pages tall automatically (i.e., use as many pages as needed for rows).
        // This combination is useful when you want a printable report that always
        // fits the page width, regardless of the number of columns, while allowing
        // the height to flow onto multiple pages.
        // ------------------------------------------------------------

        // Fit all columns to one page width; let height adjust automatically
        pageSetup.FitToPagesWide = 1;   // one page wide
        pageSetup.FitToPagesTall = 0;   // auto-calculate pages tall

        // Ensure scaling is based on page count rather than a percentage
        pageSetup.IsPercentScale = false;

        // (Optional) Demonstrate explicit setting of both dimensions using SetFitToPages
        // pageSetup.SetFitToPages(2, 3); // Uncomment to fit to 2 pages wide and 3 pages tall

        // Save the workbook (lifecycle save rule)
        workbook.Save("FitToPagesDemo.xlsx");

        // Render the first printed page to an image to visualize the scaling effect
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png
        };
        SheetRender renderer = new SheetRender(sheet, renderOptions);
        renderer.ToImage(0, "FitToPagesDemo_Page1.png");
    }
}