// Title: Convert Excel WordArt to SVG with Aspose.Cells – Preserve Gradient Fills as Vectors
// Description: C# sample that loads an .xlsx workbook, accesses the first worksheet, configures SvgImageOptions (FitToViewPort, CssPrefix), and uses SheetRender to export the sheet—including WordArt shapes with gradient fills—to a scalable SVG file, with error handling for missing files.
// Keywords: Aspose.Cells SVG export | Excel WordArt to SVG | preserve gradient fills | C# SheetRender SVG | vector gradient WordArt | .NET Excel to SVG conversion
// Common Searches: Aspose.Cells export WordArt as SVG | keep gradient fill when converting Excel to SVG | C# render worksheet with WordArt to SVG | SvgImageOptions gradient support Aspose | convert Excel workbook to vector SVG .NET
// Developer Intent: Create an SVG representation of an Excel worksheet that contains WordArt, ensuring that any gradient fills are retained as vector data rather than rasterized images.
// Use Cases: Generate web‑ready SVG graphics from Excel reports that include decorative WordArt. | Produce print‑quality SVG assets for marketing collateral while preserving original gradient styling. | Automate batch conversion of multiple Excel files with WordArt into vector SVGs for design pipelines.
// AI Prompts: Write C# code using Aspose.Cells to load an .xlsx file and export the first worksheet with WordArt to SVG, keeping gradient fills as vectors. | Explain how SvgImageOptions properties like FitToViewPort and CssPrefix affect the SVG output of WordArt shapes. | Show best practices for handling missing input files and exceptions during Excel‑to‑SVG conversion with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWordArtToSvg
{
    // C# sample that loads an .xlsx workbook, accesses the first worksheet, configures SvgImageOptions (FitToViewPort, CssPrefix), and uses SheetRender to export the sheet—including WordArt shapes with gradient fills—to a scalable SVG file, with error handling for missing files.
    class Program
    {
        static void Main()
        {
            const string inputPath = "WordArtWorkbook.xlsx";
            const string outputPath = "WordArtOutput.svg";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains WordArt shapes
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or any worksheet that contains the WordArt)
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure SVG rendering options
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    // Ensure the generated SVG fits the view port
                    FitToViewPort = true,

                    // Optional: add a CSS prefix to avoid style conflicts
                    CssPrefix = "wa-"
                };

                // Render the worksheet (including WordArt) to an SVG file
                // The page index is 0 because we are rendering the first (and only) page of the sheet
                SheetRender renderer = new SheetRender(worksheet, svgOptions);
                renderer.ToImage(0, outputPath);

                Console.WriteLine($"Worksheet with WordArt has been successfully saved as SVG: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
