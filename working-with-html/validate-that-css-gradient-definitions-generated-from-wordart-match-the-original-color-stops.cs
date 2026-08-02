// Title: Validate CSS Linear‑Gradient from Aspose.Cells WordArt Gradient Fill (C#)
// Description: Creates a workbook, adds a WordArt shape, applies a two‑color gradient with transparency via SetTwoColorGradient, extracts the GradientStopCollection, converts it to a CSS linear‑gradient string, compares the result with an expected rgba definition, outputs the match status, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | WordArt gradient | SetTwoColorGradient | CSS linear-gradient | rgba opacity | gradient stop conversion | HTML export | gradient validation
// Common Searches: convert Aspose.Cells gradient stops to CSS linear-gradient | C# validate WordArt gradient CSS | Aspose.Cells SetTwoColorGradient transparency to rgba | generate CSS gradient from Excel WordArt | compare Aspose gradient fill with expected CSS
// Developer Intent: Ensure that the CSS linear‑gradient generated from a WordArt shape’s gradient fill exactly matches the intended color stops and opacity values.
// Use Cases: Produce CSS for web previews of Excel WordArt gradients to maintain visual fidelity. | Automated testing of gradient definitions when exporting Excel to HTML in CI pipelines. | Validate design specifications by comparing exported gradient CSS against a reference style.
// AI Prompts: Write a method that takes an Aspose.Cells GradientStopCollection and returns a CSS linear‑gradient string with correct direction and rgba opacity. | Show how to extend the validation to support any number of gradient stops and all GradientStyleType directions. | Create NUnit tests that verify the generated CSS matches expected strings for various start/end colors and transparency levels.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientValidation
{
    // Creates a workbook, adds a WordArt shape, applies a two‑color gradient with transparency via SetTwoColorGradient, extracts the GradientStopCollection, converts it to a CSS linear‑gradient string, compares the result with an expected rgba definition, outputs the match status, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a WordArt shape (using AddWordArt)
                // Parameters: style, text, upper left row, upper left column, top, left, width, height
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "WordArt",
                    2, 0, 2, 0, 300, 100);

                // Set the fill type to gradient so we can work with GradientFill
                wordArt.Fill.FillType = FillType.Gradient;

                // Configure a two‑color gradient (rule: SetTwoColorGradient)
                // Example: from semi‑transparent red at start to opaque blue at end
                Color startColor = Color.Red;
                double startTransparency = 0.3; // 30% transparent
                Color endColor = Color.Blue;
                double endTransparency = 0.0;   // opaque
                GradientStyleType style = GradientStyleType.Horizontal;
                int variant = 1;

                // Use GradientFill's SetTwoColorGradient overload with transparency values
                wordArt.Fill.GradientFill.SetTwoColorGradient(
                    startColor, startTransparency,
                    endColor, endTransparency,
                    style, variant);

                // Retrieve the gradient stops created by the above call
                GradientStopCollection stops = wordArt.Fill.GradientFill.GradientStops;

                // Build a CSS linear‑gradient string from the stops
                string cssGradient = BuildCssGradient(stops, style);
                Console.WriteLine("Generated CSS gradient:");
                Console.WriteLine(cssGradient);

                // Define the original CSS gradient we expect (based on the parameters above)
                // Transparency is expressed as alpha (0‑1) in CSS; 0.3 transparency => alpha = 0.7 opacity
                string expectedCss = "linear-gradient(to right, rgba(255,0,0,0.70) 0%, rgba(0,0,255,1.00) 100%)";

                // Validate that the generated CSS matches the expected definition
                bool isMatch = string.Equals(cssGradient, expectedCss, StringComparison.OrdinalIgnoreCase);
                Console.WriteLine($"Validation result: {(isMatch ? "Match" : "Mismatch")}");

                // Save the workbook (lifecycle rule: save)
                string outputPath = "WordArtGradientValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to convert gradient stops to a CSS linear‑gradient string
        private static string BuildCssGradient(GradientStopCollection stops, GradientStyleType style)
        {
            // Determine direction keyword based on GradientStyleType
            string direction = style switch
            {
                GradientStyleType.Horizontal => "to right",
                GradientStyleType.Vertical => "to bottom",
                GradientStyleType.DiagonalDown => "to bottom right",
                GradientStyleType.DiagonalUp => "to top right",
                GradientStyleType.FromCenter => "circle",
                _ => "to right"
            };

            // Build stop list
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append("linear-gradient(");
            sb.Append(direction);
            sb.Append(", ");

            for (int i = 0; i < stops.Count; i++)
            {
                GradientStop stop = stops[i];
                // CellsColor provides Color and Transparency (0‑1, where 0 = opaque)
                Color col = stop.CellsColor.Color;
                double transparency = stop.CellsColor.Transparency; // 0 = opaque, 1 = fully transparent
                double opacity = 1.0 - transparency; // CSS opacity (0‑1)

                // Position is expressed as a percentage (0‑100)
                double positionPercent = stop.Position * 100.0;

                sb.Append($"rgba({col.R},{col.G},{col.B},{opacity:F2}) {positionPercent:F0}%");
                if (i < stops.Count - 1)
                    sb.Append(", ");
            }

            sb.Append(")");
            return sb.ToString();
        }
    }
}
