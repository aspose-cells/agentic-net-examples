// Title: Render WordArt with a configurable gradient direction and export to PDF or PNG using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set linear gradient direction on WordArt shape | Export WordArt with gradient fill to PDF using Aspose.Cells | Programmatically change WordArt gradient direction in a .NET workbook | Save workbook containing WordArt gradient as PNG with Aspose.Cells | Parse command line argument to GradientDirectionType in Aspose.Cells example | Apply opposite gradient direction to WordArt using Aspose.Cells API
// Tags: Aspose.Cells WordArt linear gradient | configure gradient direction Aspose.Cells | PDF export of WordArt gradient Aspose.Cells | PNG rendering of WordArt gradient Aspose.Cells | C# GradientDirectionType parsing example

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace WordArtGradientDemo
{
    // Shows how to create a workbook, add a WordArt shape, apply a red‑to‑blue linear gradient with a direction parsed from command‑line arguments, save the result as PDF (PDF/A‑1b) and PNG, then reapply the gradient using the opposite direction and generate a second PDF.
    class Program
    {
        // Convert a string argument to the Aspose.Cells GradientDirectionType enum.
        private static GradientDirectionType ParseDirection(string dir)
        {
            if (string.IsNullOrWhiteSpace(dir))
                return GradientDirectionType.FromUpperLeftCorner; // default

            switch (dir.Trim().ToLowerInvariant())
            {
                case "topleft":
                case "upperleftcorner":
                    return GradientDirectionType.FromUpperLeftCorner;
                case "topright":
                case "upperrightcorner":
                    return GradientDirectionType.FromUpperRightCorner;
                case "bottomleft":
                case "lowerleftcorner":
                    return GradientDirectionType.FromLowerLeftCorner;
                case "bottomright":
                case "lowerrightcorner":
                    return GradientDirectionType.FromLowerRightCorner;
                case "center":
                    return GradientDirectionType.FromCenter;
                default:
                    Console.WriteLine($"Unrecognized direction \"{dir}\", falling back to UpperLeftCorner.");
                    return GradientDirectionType.FromUpperLeftCorner;
            }
        }

        // Get the opposite gradient direction.
        private static GradientDirectionType GetOppositeDirection(GradientDirectionType dir)
        {
            switch (dir)
            {
                case GradientDirectionType.FromUpperLeftCorner:
                    return GradientDirectionType.FromLowerRightCorner;
                case GradientDirectionType.FromUpperRightCorner:
                    return GradientDirectionType.FromLowerLeftCorner;
                case GradientDirectionType.FromLowerLeftCorner:
                    return GradientDirectionType.FromUpperRightCorner;
                case GradientDirectionType.FromLowerRightCorner:
                    return GradientDirectionType.FromUpperLeftCorner;
                case GradientDirectionType.FromCenter:
                    return GradientDirectionType.FromCenter;
                default:
                    return GradientDirectionType.FromUpperLeftCorner;
            }
        }

        static void Main(string[] args)
        {
            // Expected usage:
            //   WordArtGradientDemo.exe [GradientDirection]
            // Example:
            //   WordArtGradientDemo.exe TopLeft

            string directionArg = args.Length > 0 ? args[0] : "UpperLeftCorner";
            GradientDirectionType gradientDirection = ParseDirection(directionArg);
            Console.WriteLine($"Using gradient direction: {gradientDirection}");

            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Add a WordArt shape.
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "Gradient WordArt",
                    2,   // topRow
                    20,  // top (pixels)
                    2,   // leftColumn
                    20,  // left (pixels)
                    100, // height (pixels)
                    400  // width (pixels)
                );

                // 3. Configure the gradient fill for the WordArt.
                FillFormat fill = wordArt.Fill;               // use the non‑obsolete Fill property
                fill.FillType = FillType.Gradient;            // ensure gradient fill is used

                GradientFill gradient = fill.GradientFill;
                // Set a linear gradient with the chosen direction (degree = 0 for default)
                gradient.SetGradient(GradientFillType.Linear, 0, gradientDirection);

                // Define two color stops: start (red) at 0% and end (blue) at 100%.
                gradient.GradientStops.Clear();
                gradient.GradientStops.Add(0.0, Color.Red, 255);
                gradient.GradientStops.Add(1.0, Color.Blue, 255);

                // 4. Save the workbook to PDF (gradient is preserved in PDF rendering).
                string pdfPath = "WordArtGradient.pdf";
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Ensure that WordArt is rendered as vector graphics.
                    Compliance = PdfCompliance.PdfA1b
                };
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF saved to {pdfPath}");

                // 5. Save the workbook to PNG image (gradient is rendered in raster image).
                string pngPath = "WordArtGradient.png";
                workbook.Save(pngPath, SaveFormat.Png);
                Console.WriteLine($"PNG image saved to {pngPath}");

                // 6. Demonstrate that the gradient direction can be changed programmatically.
                GradientDirectionType opposite = GetOppositeDirection(gradientDirection);
                // Re‑apply the gradient with the opposite direction.
                gradient.SetGradient(GradientFillType.Linear, 0, opposite);
                string pdfOpposite = "WordArtGradient_Opposite.pdf";
                workbook.Save(pdfOpposite, pdfOptions);
                Console.WriteLine($"Opposite direction PDF saved to {pdfOpposite}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
