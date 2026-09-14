// Title: Generate a transparent PNG from a CSV file using Aspose.Cells in C#
// AI Prompts: Write C# code that loads a CSV file into an Aspose.Cells Workbook, sets ImageOrPrintOptions to PNG with Transparent=true, and saves the first worksheet as a single transparent PNG image. | Show how to use SheetRender together with ImageOrPrintOptions to export a worksheet containing CSV data to a transparent PNG file in .NET.
// Common Searches: aspnet convert csv to png with transparent background using aspose.cells | c# export worksheet as transparent png image from csv data | how to render csv data to transparent png with aspose cells library | save aspose.cells worksheet to png with alpha channel .net
// Tags: Aspose.Cells CSV to transparent PNG | ImageOrPrintOptions transparent background | SheetRender PNG export | C# workbook import CSV and render image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTransparentPngDemo
{
    // // Loads a CSV file into a new Aspose.Cells Workbook, configures ImageOrPrintOptions for PNG with a transparent background, and uses SheetRender to save the first worksheet as a single-page transparent PNG image.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "data.csv";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Import CSV data into the first worksheet (starting at cell A1)
            // Using comma as delimiter, converting numeric data, and placing at row 0, column 0
            workbook.Worksheets[0].Cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Configure image rendering options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png, // Output format PNG
                Transparent = true,                             // Enable transparent background
                OnePagePerSheet = true                          // Render the whole sheet as a single page
            };

            // Create a SheetRender for the first worksheet with the above options
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);

            // Render the first (and only) page to a PNG file with transparent background
            string outputImagePath = "output.png";
            sheetRender.ToImage(0, outputImagePath);

            Console.WriteLine($"Workbook rendered to transparent PNG: {outputImagePath}");
        }
    }
}
