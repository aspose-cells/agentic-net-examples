using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPageBreakPreview
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample data to generate multiple pages
            for (int row = 0; row < 200; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Configure page setup: define print area and fit to one page wide
            PageSetup ps = sheet.PageSetup;
            ps.PrintArea = "A1:E200";
            ps.FitToPagesWide = 1;   // force horizontal pagination to a single page
            ps.FitToPagesTall = 0;   // allow vertical pagination (multiple pages)

            // Switch worksheet to Page Break Preview view
            sheet.ViewType = ViewType.PageBreakPreview; // alternatively: sheet.IsPageBreakPreview = true;

            // Create rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.OnePagePerSheet = true; // each page rendered as a separate image

            // Create SheetRender to evaluate pagination
            SheetRender render = new SheetRender(sheet, options);

            // Output total page count
            Console.WriteLine($"Total pages in preview: {render.PageCount}");

            // Render each page to an image file for visual preview
            for (int i = 0; i < render.PageCount; i++)
            {
                string fileName = $"PagePreview_{i + 1}.png";
                render.ToImage(i, fileName);
                Console.WriteLine($"Rendered page {i + 1} to {fileName}");
            }

            // Save the workbook (optional, to keep the preview view)
            workbook.Save("PageBreakPreviewDemo.xlsx");

            // Release resources
            render.Dispose();
        }
    }
}