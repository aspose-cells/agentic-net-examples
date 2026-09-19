// Title: How to apply a blue-to-transparent linear gradient fill to a WordArt shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a WordArt shape, clears any existing gradient stops, and adds a solid‑blue start stop and a fully transparent end stop. | Show the steps to set a shape's FillType to Gradient and configure its GradientFill object for a blue‑to‑transparent linear gradient in an Excel file. | Provide a complete example that adds a WordArt shape, defines a linear gradient from solid blue to 100% transparency, and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# example for setting a gradient fill on a WordArt shape with transparent end color | how to create a blue to transparent gradient for WordArt in an Excel file using Aspose.Cells | C# Aspose.Cells linear gradient fill on shape with custom gradient stops
// Tags: Aspose.Cells WordArt gradient fill C# | linear gradient fill shape Aspose.Cells | transparent gradient stop Aspose.Cells | configure WordArt fill Aspose.Cells | C# Excel shape gradient Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a workbook, adds a WordArt shape, applies a blue‑to‑transparent linear gradient fill, and saves the file as WordArtGradient.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape to the worksheet
            // Parameters: style, text, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, rowOffset, columnOffset
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "WordArt Example",
                0, 0, 5, 5, 0, 0);

            // Adjust size (optional)
            wordArt.Width = 400;
            wordArt.Height = 100;

            // Configure the fill to be a linear gradient from blue to transparent
            wordArt.Fill.FillType = FillType.Gradient;

            // Access the GradientFill object
            GradientFill gradient = wordArt.Fill.GradientFill;

            // Define gradient stops (position, color, transparency)
            gradient.GradientStops.Clear();
            // Position is an integer (0‑100), transparency is also an integer (0‑100)
            gradient.GradientStops.Add(0, Color.FromArgb(255, 0, 0, 255), 0);     // Start: solid blue
            gradient.GradientStops.Add(100, Color.FromArgb(255, 0, 0, 255), 100); // End: fully transparent

            // Save the workbook to a file
            workbook.Save("WordArtGradient.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
