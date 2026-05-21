using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsRenderExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to have visible content)
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["B2"].PutValue(123);
            worksheet.Cells["C3"].PutValue(DateTime.Now);

            // Hide gridlines for rendering
            worksheet.IsGridlinesVisible = false;

            // Configure image rendering options for JPEG
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Jpeg,
                // Optional: set JPEG quality (0-100)
                Quality = 90
            };

            // Create a SheetRender instance for the worksheet
            SheetRender sheetRender = new SheetRender(worksheet, options);

            // Render the first page (index 0) to a JPEG file
            sheetRender.ToImage(0, "RenderedWorksheet.jpg");

            // Clean up resources
            sheetRender.Dispose();

            Console.WriteLine("Worksheet rendered to JPEG with gridlines hidden.");
        }
    }
}