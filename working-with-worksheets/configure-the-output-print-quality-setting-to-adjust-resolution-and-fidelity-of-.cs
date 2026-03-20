using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPrintQualityDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Configure page setup print quality (DPI)
            // This property directly influences the printer resolution.
            // -------------------------------------------------
            worksheet.PageSetup.PrintQuality = 300; // Set to 300 DPI

            // -------------------------------------------------
            // Optional: Configure image/print options for rendering
            // These settings affect image generation and printing via SheetRender.
            // -------------------------------------------------
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                // Set image type to JPEG to demonstrate Quality property
                ImageType = Aspose.Cells.Drawing.ImageType.Jpeg,
                // JPEG quality (0-100)
                Quality = 90,
                // Horizontal and vertical resolution (DPI)
                HorizontalResolution = 300,
                VerticalResolution = 300,
                // Print each sheet on a separate page
                OnePagePerSheet = true
            };

            // Create a SheetRender instance with the configured options
            SheetRender sheetRender = new SheetRender(worksheet, printOptions);

            // Render the first page to an image file (optional, demonstrates the settings)
            sheetRender.ToImage(0, "PrintedPage.jpg");

            // Save the workbook with the configured print quality
            workbook.Save("PrintQualityConfigured.xlsx");

            // Clean up resources
            sheetRender.Dispose();

            Console.WriteLine("Workbook saved with print quality set to 300 DPI.");
        }
    }
}