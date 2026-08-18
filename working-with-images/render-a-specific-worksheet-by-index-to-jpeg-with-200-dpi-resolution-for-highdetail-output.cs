// Title: Export a Specific Worksheet to a 200 DPI JPEG with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, checks a zero‑based worksheet index, sets ImageOrPrintOptions for JPEG with 200 DPI horizontal and vertical resolution, creates a SheetRender for that sheet, and saves the first page as a single‑page JPEG image.
// Keywords: Aspose.Cells | C# | export worksheet to JPEG | 200 DPI | ImageOrPrintOptions | SheetRender | high‑detail Excel image | specific worksheet index
// Common Searches: Aspose.Cells render worksheet by index to JPEG | 200 DPI Excel sheet image C# | how to export a single sheet as high resolution JPEG | ImageOrPrintOptions DPI setting Aspose.Cells | SheetRender ToImage example for specific worksheet
// Developer Intent: Create a detailed JPEG image of a chosen worksheet using a 200 DPI setting.
// Use Cases: Generate printable previews of a selected sheet for reports. | Provide high‑quality images of a particular worksheet for web portals or documentation. | Automate conversion of targeted sheets from many workbooks into 200 DPI JPEG files.
// AI Prompts: Write C# code with Aspose.Cells to render worksheet index 5 as a PNG at 300 DPI and save it as 'sheet5.png'. | Explain error handling for invalid worksheet indexes when exporting to JPEG with custom DPI in Aspose.Cells. | Show how to configure ImageOrPrintOptions to produce multi‑page JPEG output instead of a single page.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsRenderingDemo
{
    // Loads an Excel workbook, checks a zero‑based worksheet index, sets ImageOrPrintOptions for JPEG with 200 DPI horizontal and vertical resolution, creates a SheetRender for that sheet, and saves the first page as a single‑page JPEG image.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Path for the output JPEG image
            string outputImage = "worksheet_page0_200dpi.jpg";

            // Index of the worksheet to render (0‑based)
            int worksheetIndex = 2; // example: third worksheet

            // Load the workbook (uses the provided Workbook constructor rule)
            Workbook workbook = new Workbook(sourceFile);

            // Validate worksheet index
            if (worksheetIndex < 0 || worksheetIndex >= workbook.Worksheets.Count)
            {
                Console.WriteLine("Invalid worksheet index.");
                return;
            }

            // Get the target worksheet
            Worksheet sheet = workbook.Worksheets[worksheetIndex];

            // Configure image rendering options (uses ImageOrPrintOptions rule)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,          // JPEG output
                HorizontalResolution = 200,          // 200 DPI horizontal
                VerticalResolution = 200,            // 200 DPI vertical
                OnePagePerSheet = true               // Render each sheet as a single page
            };

            // Create a SheetRender instance for the selected worksheet (uses SheetRender constructor rule)
            SheetRender renderer = new SheetRender(sheet, options);

            // Render the first page of the worksheet to a JPEG file (uses SheetRender.ToImage overload)
            renderer.ToImage(0, outputImage);

            Console.WriteLine($"Worksheet {worksheetIndex} rendered to JPEG at {outputImage} with 200 DPI.");
        }
    }
}
