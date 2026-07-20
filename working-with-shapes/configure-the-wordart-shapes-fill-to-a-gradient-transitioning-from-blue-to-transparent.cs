// Title: Aspose.Cells .NET – Apply a Blue‑to‑Transparent Horizontal Gradient Fill to a WordArt Shape
// Description: This example creates a new workbook, inserts a WordArt shape, and uses the FillFormat.SetTwoColorGradient method to apply a horizontal gradient that fades from opaque blue to fully transparent blue, then saves the file as WordArtGradient.xlsx.
// Keywords: Aspose.Cells | C# | WordArt gradient | transparent fill | SetTwoColorGradient | horizontal gradient | Excel shape fill | fill format example | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells set gradient fill on WordArt | C# WordArt transparent gradient Excel | SetTwoColorGradient WordArt Aspose | horizontal blue gradient WordArt .NET | how to make WordArt fade to transparent in Excel
// Developer Intent: Add a WordArt shape and configure its fill to a horizontal blue‑to‑transparent gradient.
// Use Cases: Create report titles that subtly blend into the worksheet background. | Design dashboard headers with brand‑colored gradient effects. | Produce marketing spreadsheets where WordArt fades for a polished look.
// AI Prompts: Generate code to change the gradient direction to vertical while keeping the blue‑to‑transparent colors. | Show how to add multiple WordArt shapes, each with distinct gradient colors and transparency levels, using Aspose.Cells for .NET. | Explain how to read and modify the gradient settings of an existing WordArt shape in a saved workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a new workbook, inserts a WordArt shape, and uses the FillFormat.SetTwoColorGradient method to apply a horizontal gradient that fades from opaque blue to fully transparent blue, then saves the file as WordArtGradient.xlsx.
public class WordArtGradientDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape.
        // Parameters: style, text, upper left row, upper left column,
        // row offset (pixels), column offset (pixels), height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,
            "Gradient WordArt",
            2,          // upper left row
            0,          // upper left column
            0,          // row offset
            0,          // column offset
            200,        // height
            400);       // width

        // Access the FillFormat of the WordArt shape
        FillFormat fill = wordArt.Fill;

        // Apply a two‑color gradient: opaque blue to fully transparent blue
        fill.SetTwoColorGradient(
            Color.Blue,   // first color (opaque)
            0.0,          // transparency for first color (0 = opaque)
            Color.Blue,   // second color (same hue)
            1.0,          // transparency for second color (1 = fully transparent)
            GradientStyleType.Horizontal, // gradient direction
            1);           // variant

        // Save the workbook with error handling
        try
        {
            workbook.Save("WordArtGradient.xlsx");
            Console.WriteLine("Workbook saved as WordArtGradient.xlsx");
        }
        catch (Exception saveEx)
        {
            Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
        }
    }
}
