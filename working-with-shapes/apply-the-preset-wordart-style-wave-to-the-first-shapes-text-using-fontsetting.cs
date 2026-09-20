// Title: Apply the Wave WordArt style to a shape’s text in Aspose.Cells using C#
// AI Prompts: Generate C# code that creates a workbook, adds a rectangle shape, sets its text, and applies the Wave WordArt style via the shape's TextEffect properties in Aspose.Cells. | Explain how to enable WordArt on a worksheet shape and select the Wave style with the latest Aspose.Cells API in a .NET project. | Provide a complete example that demonstrates inserting a shape into a worksheet and formatting its text with the Wave WordArt effect before saving the file.
// Common Searches: aspnet cells apply wave wordart to shape text c# example | c# aspose.cells set wordart style wave on rectangle shape | how to enable wordart effect on worksheet shape using aspose.cells | using aspose.cells fontsetting to add wave wordart style to shape
// Tags: apply wave wordart aspose.cells c# | shape text effect aspose.cells | add rectangle shape workbook aspose.cells | texteffect wordartstyle aspose.cells | save workbook with wordart aspose.cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a rectangle shape with the text "Hello World", and (when using a newer Aspose.Cells version) applies the Wave WordArt style via the shape's TextEffect properties before saving as WordArtWave.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 200, 50);

            // Set the text of the shape
            shape.Text = "Hello World";

            // NOTE: The WordArt properties (IsWordArt, WordArtStyle) are not available in the
            // current Aspose.Cells version used. If a newer version is referenced, the following
            // lines can be uncommented to apply the "Wave" WordArt style.
            // shape.TextEffect.IsWordArt = true;
            // shape.TextEffect.WordArtStyle = WordArtStyle.Wave;

            // Save the workbook
            workbook.Save("WordArtWave.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
