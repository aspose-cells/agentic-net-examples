// Title: How to retrieve and log a shape's text shadow preset using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get current shadow preset of a shape's text | read shape text shadow effect type with Aspose.Cells .NET | debug shape text shadow preset in an Aspose.Cells workbook | retrieve PresetShadowType from a rectangle shape in Aspose.Cells | log shape text shadow preset value in C# Aspose.Cells example
// Tags: shape text shadow preset Aspose.Cells | shadow effect type property C# | Aspose.Cells shape shadow effect | debug shape formatting Aspose.Cells | C# Aspose.Cells rectangle shape styling

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// The example creates a workbook, adds a rectangle shape, sets its TextBody, applies an OffsetDiagonalBottomRight shadow preset via Shape.ShadowEffect.PresetType, then reads back the PresetShadowType value, writes it to the console for debugging, and saves the file as ShadowPresetDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, lower right row, lower right column, width (pixels), height (pixels)
            Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 100);

            // Set some text inside the shape
            shape.TextBody.Text = "Aspose.Cells Text";

            // Apply a known shadow preset to the shape (this also affects the text rendering)
            shape.ShadowEffect.PresetType = PresetShadowType.OffsetDiagonalBottomRight;

            // Retrieve the current shadow preset of the shape's text
            PresetShadowType currentPreset = shape.ShadowEffect.PresetType;

            // Log the retrieved preset for debugging purposes
            Console.WriteLine($"Current shadow preset of the shape's text: {currentPreset}");

            // Save the workbook to verify the shape appears as expected
            workbook.Save("ShadowPresetDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
