// Title: Export Excel worksheets to 300 DPI JPEG images with Aspose.Cells for .NET
// Description: Loads an .xlsx file, configures ImageOrPrintOptions for JPEG at 300 DPI, renders each worksheet as a single page, and saves the images to a folder while sanitizing sheet names for file‑system safety.
// Keywords: Aspose.Cells JPEG export | 300 DPI Excel image | C# sheet to JPEG | SheetRender high resolution | Excel to image conversion .NET
// Common Searches: Aspose.Cells export each sheet to JPEG 300 DPI C# | Render Excel worksheets as high‑resolution JPEG images | C# convert workbook sheets to printable JPEG files | How to save Excel sheets as 300 DPI images using Aspose
// Developer Intent: Generate a separate 300 DPI JPEG file for every worksheet in an Excel workbook.
// Use Cases: Produce printable graphics of financial dashboards for stakeholder reports. | Create high‑quality product catalog pages from spreadsheet data for marketing materials. | Automate batch conversion of multi‑sheet workbooks into image assets for web or mobile apps.
// AI Prompts: Write C# code that uses Aspose.Cells to export all worksheets of a workbook to 300 DPI JPEG files, handling invalid characters in sheet names. | Show how to adapt the sample to output single‑page PDFs while keeping 300 DPI resolution. | Provide a script to process every Excel file in a directory, converting each sheet to high‑resolution JPEG images with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ExportSheetsAsJpeg
{
    // Loads an .xlsx file, configures ImageOrPrintOptions for JPEG at 300 DPI, renders each worksheet as a single page, and saves the images to a folder while sanitizing sheet names for file‑system safety.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Configure image options for JPEG output at 300 DPI
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,          // Export as JPEG
                HorizontalResolution = 300,          // 300 DPI horizontally
                VerticalResolution = 300,            // 300 DPI vertically
                OnePagePerSheet = true               // Render each sheet as a single page
            };

            // Ensure the output directory exists
            string outputDir = "ExportedImages";
            Directory.CreateDirectory(outputDir);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Create a SheetRender for the current worksheet with the specified options
                SheetRender sheetRender = new SheetRender(sheet, options);

                // Since OnePagePerSheet = true, PageCount will be 1.
                // Loop through pages in case the setting is changed later.
                for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                {
                    // Build a file name that includes the sheet name and page index
                    string safeSheetName = MakeFileSystemSafe(sheet.Name);
                    string fileName = Path.Combine(outputDir, $"{safeSheetName}_page{pageIndex}.jpg");

                    // Render the page to a JPEG file
                    sheetRender.ToImage(pageIndex, fileName);
                }

                // Release resources used by the renderer
                sheetRender.Dispose();
            }

            Console.WriteLine("All sheets have been exported as 300 DPI JPEG images.");
        }

        // Helper method to remove invalid filename characters from sheet names
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
