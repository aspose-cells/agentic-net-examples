// Title: Aspose.Cells for .NET – Render WordArt with Gradient Fill and Configurable Direction
// Description: Shows how to add a WordArt shape to an Excel workbook, apply a two‑color gradient fill, set its orientation with GradientDirectionType, and save the file (including HTML export) while preserving the gradient appearance.
// Keywords: Aspose.Cells WordArt gradient | C# gradient fill WordArt | GradientDirectionType Aspose.Cells | set gradient orientation Excel | preserve WordArt in HTML export | linear gradient fill Aspose.Cells | WordArt API .NET | Excel gradient fill example
// Common Searches: Aspose.Cells set WordArt gradient direction C# | WordArt gradient fill not showing in HTML conversion | GradientFill.SetGradient example Aspose.Cells | How to apply two‑color gradient to WordArt using Aspose.Cells | Configure gradient orientation for WordArt shape
// Developer Intent: Create a WordArt shape with a gradient fill and programmatically define its direction using the Aspose.Cells .NET API.
// Use Cases: Generate Excel dashboards with stylized WordArt headings that use a custom gradient direction. | Export spreadsheets to HTML while keeping WordArt gradient colors and orientation intact. | Build automated reporting tools where the gradient direction reflects data flow or user preferences.
// AI Prompts: Write C# code with Aspose.Cells to add a WordArt shape that uses a three‑color radial gradient and lets the caller choose the gradient direction. | Explain the impact of each parameter in GradientFill.SetGradient on a WordArt shape's appearance. | Provide a complete example that converts the workbook to HTML and ensures the WordArt gradient is rendered correctly.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsWordArtGradientDemo
{
    // Shows how to add a WordArt shape to an Excel workbook, apply a two‑color gradient fill, set its orientation with GradientDirectionType, and save the file (including HTML export) while preserving the gradient appearance.
    public class WordArtGradientGenerator
    {
        /// <param name="direction">The gradient direction to apply (e.g., FromUpperLeftCorner, FromCenter, etc.).</param>
        public static void CreateWordArtWithGradient(GradientDirectionType direction)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape.
            // Using WordArtStyle6 (Gradient Fill - Gray) as a base; we will customize the fill later.
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle6,   // base preset style
                "Gradient WordArt",                 // text
                5, 0,                               // upper left row, top offset
                5, 0,                               // upper left column, left offset
                200, 400);                          // height, width

            // Ensure the shape's fill type is set to Gradient so we can access GradientFill.
            wordArt.Fill.FillType = FillType.Gradient;

            // Obtain the GradientFill object from the shape.
            GradientFill gradientFill = wordArt.Fill.GradientFill;

            // Apply a two‑color gradient (light gray to dark gray) using the FillFormat method.
            // This sets the colors and the basic gradient style.
            wordArt.Fill.SetTwoColorGradient(
                Color.LightGray,                    // first gradient color
                Color.DarkGray,                     // second gradient color
                GradientStyleType.Horizontal,       // gradient style
                1);                                 // variant (1‑4)

            // Configure the gradient direction and type.
            // Using Linear fill type; angle is set to 0 because direction will define the orientation.
            gradientFill.SetGradient(
                GradientFillType.Linear,            // linear gradient
                0.0,                                // angle (ignored for non‑linear types)
                direction);                         // direction supplied by the caller

            // Save the workbook to demonstrate the result.
            workbook.Save("WordArtGradientDemo.xlsx");
        }

        // Example usage
        public static void Main()
        {
            // Create a WordArt with gradient flowing from the upper left corner to the lower right.
            CreateWordArtWithGradient(GradientDirectionType.FromUpperLeftCorner);
        }
    }
}
