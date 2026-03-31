using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFitToPagesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with sample data to span multiple pages
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Access the PageSetup object of the worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Define the print area (optional, but helps control what is rendered)
            pageSetup.PrintArea = "A1:J100";

            // Configure FitToPages to scale the sheet to 2 pages wide and 3 pages tall
            // This uses the SetFitToPages method which directly sets both properties.
            pageSetup.SetFitToPages(2, 3);

            // Ensure that scaling is controlled by FitToPages (not percent scale)
            pageSetup.IsPercentScale = false;

            // Create rendering options (default options are sufficient for this demo)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();

            // Create a SheetRender instance after page setup configuration
            SheetRender renderer = new SheetRender(sheet, renderOptions);

            // Output the calculated page scale and total page count
            Console.WriteLine($"Calculated page scale: {renderer.PageScale * 100}%");
            Console.WriteLine($"Total pages after FitToPages(2,3): {renderer.PageCount}");

            // Render each page to an image file (optional demonstration)
            for (int i = 0; i < renderer.PageCount; i++)
            {
                string fileName = $"Page_{i + 1}.png";
                renderer.ToImage(i, fileName);
                Console.WriteLine($"Rendered page {i + 1} to {fileName}");
            }

            // Save the workbook (the scaling settings are preserved in the file)
            workbook.Save("FitToPagesDemo.xlsx");
        }
    }
}