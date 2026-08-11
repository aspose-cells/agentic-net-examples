// Title: Render WordArt with Two‑Color Gradient and Adjustable Direction using Aspose.Cells for .NET
// Description: Shows how to create an Excel workbook, add a WordArt shape, apply a two‑color gradient fill, set the gradient direction via the GradientDirectionType enum, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | WordArt gradient | GradientDirectionType | C# Excel API | two color gradient fill | configure gradient direction | Excel shape styling | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set WordArt gradient direction | C# add WordArt with gradient fill | how to change gradient direction of WordArt using Aspose | GradientFill example Aspose.Cells | create WordArt shape programmatically Aspose.Cells
// Developer Intent: Create an Excel file that contains a WordArt shape with a two‑color gradient whose direction can be defined programmatically.
// Use Cases: Design branded report headers with a consistent gradient orientation. | Generate automated dashboards where gradient direction indicates data flow. | Produce marketing templates that require stylized WordArt headings. | Batch‑process existing workbooks to update WordArt gradient styles across sheets.
// AI Prompts: Generate C# code using Aspose.Cells to add a WordArt shape with a three‑color gradient and a custom linear angle. | Show how to open an existing workbook, locate a WordArt shape by name, and modify its GradientFill colors and direction. | Explain the mapping between GradientDirectionType enum values and the visual gradient directions in Excel. | Provide a step‑by‑step guide to apply a linear gradient fill to multiple WordArt shapes across worksheets.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create an Excel workbook, add a WordArt shape, apply a two‑color gradient fill, set the gradient direction via the GradientDirectionType enum, and save the file with Aspose.Cells for .NET.
public class WordArtGradientRenderer
{
    // Creates a workbook with a WordArt shape that uses a two‑color gradient.
    // The gradient direction is supplied via the 'direction' parameter.
    public static void CreateWordArtWithGradient(string filePath, GradientDirectionType direction)
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape; use a preset style that already has a gradient base
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6, // Gradient Fill - Gray
            "Gradient WordArt",
            5,   // upperLeftRow
            0,   // top
            5,   // upperLeftColumn
            0,   // left
            200, // height
            100  // width
        );

        // Set the fill type to Gradient so we can access GradientFill
        wordArt.Fill.FillType = FillType.Gradient;

        // Obtain the GradientFill object
        GradientFill gradientFill = wordArt.Fill.GradientFill;

        // Define the two colors for the gradient (light gray to dark gray)
        gradientFill.SetTwoColorGradient(
            Color.LightGray,
            Color.DarkGray,
            GradientStyleType.Horizontal,
            1 // variant
        );

        // Apply the gradient direction supplied by the caller.
        // Angle is set to 0 because direction is handled by GradientDirectionType.
        gradientFill.SetGradient(GradientFillType.Linear, 0.0, direction);

        // Save the workbook to the specified path
        workbook.Save(filePath);
    }
}

class Program
{
    static void Main()
    {
        // Example: render WordArt with a gradient flowing from the upper left corner
        GradientDirectionType direction = GradientDirectionType.FromUpperLeftCorner;

        // Output file path
        string outputPath = "WordArtGradient.xlsx";

        // Create the WordArt with the chosen gradient direction
        WordArtGradientRenderer.CreateWordArtWithGradient(outputPath, direction);
    }
}
