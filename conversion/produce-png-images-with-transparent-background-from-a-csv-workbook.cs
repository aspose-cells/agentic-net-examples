// Title: C# – Convert CSV to Transparent PNG Images with Aspose.Cells
// Description: Learn how to import a CSV file into an Aspose.Cells workbook, configure ImageOrPrintOptions for PNG format with a transparent background, and render each worksheet page as a separate PNG file. The example saves the images to a designated folder and optionally stores the workbook as an XLSX file.
// Keywords: Aspose.Cells CSV to PNG | transparent PNG Aspose.Cells | C# convert CSV to image | ImageOrPrintOptions Transparent true | .NET render worksheet as PNG | SheetRender transparent background | export CSV as PNG .NET | one page per sheet Aspose | save workbook as XLSX | Aspose.Cells image rendering
// Common Searches: Aspose.Cells render CSV as transparent PNG | C# export worksheet to PNG with transparency | How to create PNG images from CSV using Aspose.Cells | ImageOrPrintOptions transparent background example | SheetRender one page per sheet PNG
// Developer Intent: Generate PNG files with a transparent background from a CSV‑based workbook using Aspose.Cells for .NET.
// Use Cases: Convert CSV reports into high‑quality PNG graphics for web or UI overlays. | Produce transparent PNG assets for dashboards, presentations, or marketing material. | Automate batch rendering of multiple worksheet pages while preserving transparency. | Save the original workbook for later editing or archival after image generation.
// AI Prompts: Write C# code that reads a CSV file, loads it into an Aspose.Cells workbook, and saves each sheet as a transparent PNG image. | Explain the steps to configure ImageOrPrintOptions for PNG output with a transparent background and one page per sheet in Aspose.Cells. | Provide a script that renders a CSV‑derived worksheet to PNG files in a specific output folder and also saves the workbook as XLSX.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTransparentPngDemo
{
    // Learn how to import a CSV file into an Aspose.Cells workbook, configure ImageOrPrintOptions for PNG format with a transparent background, and render each worksheet page as a separate PNG file. The example saves the images to a designated folder and optionally stores the workbook as an XLSX file.
    public class Program
    {
        public static void Main()
        {
            // Path to the source CSV file
            string csvPath = "data.csv";

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Import CSV data into the first worksheet starting at cell A1 (row 0, column 0)
            // Using comma as delimiter and converting numeric strings to numbers
            workbook.Worksheets[0].Cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Configure image rendering options:
            // - Output format: PNG
            // - Transparent background
            // - Render each worksheet as a single page
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                Transparent = true,
                OnePagePerSheet = true
            };

            // Create a SheetRender for the first worksheet with the above options
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);

            // Ensure the output directory exists
            string outputDir = "output_images";
            Directory.CreateDirectory(outputDir);

            // Render each page of the sheet to a separate PNG file with transparent background
            for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
            {
                string imagePath = Path.Combine(outputDir, $"sheet_page_{pageIndex}.png");
                sheetRender.ToImage(pageIndex, imagePath);
                Console.WriteLine($"Rendered page {pageIndex} to {imagePath}");
            }

            // Optionally, save the workbook for reference
            workbook.Save(Path.Combine(outputDir, "imported_workbook.xlsx"));
        }
    }
}
