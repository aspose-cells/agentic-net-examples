// Title: Convert an Excel worksheet containing WordArt to SVG with gradient preservation using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet with WordArt to SVG keeping gradient fills | how to retain WordArt gradient colors when saving Excel as SVG in .NET | C# code to detect WordArt shapes before converting workbook to SVG | verify gradient definitions in SVG generated from Excel using Aspose.Cells | save Excel workbook as single‑page SVG with ImageSaveOptions OnePagePerSheet
// Tags: Aspose.Cells worksheet to SVG conversion | retain WordArt gradients in SVG output | C# detect WordArt shapes in Excel | ImageSaveOptions OnePagePerSheet setting | SVG gradient verification after export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace WordArtToSvg
{
    // The example loads an Excel file, scans its shapes to identify WordArt objects, and saves the first worksheet as a single‑page SVG using Aspose.Cells ImageSaveOptions. After export it reads the SVG file and checks for <linearGradient> or <radialGradient> elements to confirm that gradient definitions from the WordArt are preserved.
    class Program
    {
        /// <param name="args">
        /// args[0] - path to the source Excel file (optional, default: "input_with_wordart.xlsx")
        /// args[1] - path to the destination SVG file (optional, default: "output.svg")
        /// </param>
        static void Main(string[] args)
        {
            // Resolve input / output paths
            string inputPath = args.Length > 0 ? args[0] : "input_with_wordart.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.svg";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Enumerate shapes and report WordArt detection (optional diagnostic)
                Console.WriteLine("Scanning shapes for WordArt...");
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.IsWordArt)
                    {
                        Console.WriteLine($"- WordArt found: Text=\"{shape.Text}\"");
                    }
                }

                // Prepare SVG save options
                ImageSaveOptions svgOptions = new ImageSaveOptions(SaveFormat.Svg);
                // Configure rendering to place each worksheet on a single page
                svgOptions.ImageOrPrintOptions.OnePagePerSheet = true;

                // Save the workbook as SVG
                workbook.Save(outputPath, svgOptions);
                Console.WriteLine($"Workbook saved as SVG to \"{outputPath}\".");

                // Verify that gradient definitions exist in the generated SVG
                VerifyGradientInSvg(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred during conversion:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="svgFilePath">Path to the SVG file.</param>
        private static void VerifyGradientInSvg(string svgFilePath)
        {
            try
            {
                if (!File.Exists(svgFilePath))
                {
                    Console.WriteLine($"Error: SVG file \"{svgFilePath}\" was not created.");
                    return;
                }

                string svgContent = File.ReadAllText(svgFilePath);
                bool hasLinearGradient = svgContent.Contains("<linearGradient");
                bool hasRadialGradient = svgContent.Contains("<radialGradient");

                if (hasLinearGradient || hasRadialGradient)
                {
                    Console.WriteLine("Gradient definitions detected in the SVG:");
                    if (hasLinearGradient) Console.WriteLine("- Linear gradients present.");
                    if (hasRadialGradient) Console.WriteLine("- Radial gradients present.");
                }
                else
                {
                    Console.WriteLine("No gradient definitions were found in the SVG. " +
                                      "If the source WordArt uses gradients, this may indicate a loss of vector data.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to verify gradient definitions:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
