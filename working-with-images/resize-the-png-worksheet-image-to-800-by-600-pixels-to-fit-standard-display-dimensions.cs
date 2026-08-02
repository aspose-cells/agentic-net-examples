using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsResizeWorksheetImage
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a new workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data to the worksheet (optional)
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B1"].PutValue(DateTime.Now);

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.ImageType = ImageType.Png;                 // output format PNG
            options.SetDesiredSize(800, 600, false);           // resize to 800x600, do not keep aspect ratio

            // Render the worksheet to an image file with the specified size
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToImage(0, "Worksheet_800x600.png");      // page index 0 (first page)

            // Optionally, save the workbook itself
            workbook.Save("Workbook.xlsx");

            Console.WriteLine("Worksheet rendered to 800x600 PNG successfully.");
        }
    }
}