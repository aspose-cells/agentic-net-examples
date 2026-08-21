// Title: Create Curved WordArt (ArchUpCurve) in Excel with Aspose.Cells for .NET
// Description: Demonstrates how to add a WordArt text‑effect shape to a worksheet, apply the built‑in ArchUpCurve preset for a curved appearance, and save the workbook. The example also notes that fine‑tuning curvature via adjustment values is not exposed in the current Aspose.Cells API.
// Keywords: Aspose.Cells curved WordArt | ArchUpCurve preset shape | C# add WordArt shape Excel | Excel text effect shape | WordArt curvature Aspose.Cells | set WordArt preset Aspose | shape collection Aspose.Cells | Excel automation C# | WordArt API limitation | save workbook with WordArt
// Common Searches: how to add curved WordArt in Aspose.Cells C# | set ArchUpCurve preset for WordArt shape | Aspose.Cells adjust WordArt curvature | C# create WordArt text effect Excel | Aspose.Cells shape collection example
// Developer Intent: Add a WordArt shape with the ArchUpCurve preset to a worksheet and generate an Excel file.
// Use Cases: Insert stylized curved headings into programmatically generated reports. | Create eye‑catching titles for marketing dashboards built with Excel. | Showcase the current API limitation where curvature adjustments are unavailable.
// AI Prompts: Generate code that uses other WordArt presets such as ArchDownCurve or Circle with Aspose.Cells for .NET. | Provide a version‑check snippet that detects the Adjustments property in newer Aspose.Cells releases and sets curvature values when possible. | Explain how to position a WordArt shape precisely using row/column indices and pixel offsets.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a WordArt text‑effect shape to a worksheet, apply the built‑in ArchUpCurve preset for a curved appearance, and save the workbook. The example also notes that fine‑tuning curvature via adjustment values is not exposed in the current Aspose.Cells API.
class CurvedWordArtExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet and its shape collection.
            Worksheet sheet = workbook.Worksheets[0];
            ShapeCollection shapes = sheet.Shapes;

            // Add a WordArt (text effect) shape.
            // Parameters: effect, text, font name, size, bold, italic,
            // topRow, top offset, leftColumn, left offset, height, width
            Shape wordArt = shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Curved Text",
                "Arial",
                36,
                false,
                false,
                2, 0,   // top row & offset
                2, 0,   // left column & offset
                100,    // height
                200);   // width

            // Apply a built‑in curved WordArt preset shape (e.g., ArchUpCurve).
            wordArt.TextEffect.PresetShape = MsoPresetTextEffectShape.ArchUpCurve;

            // Note: Adjustments property is not available in the current Aspose.Cells version.
            // Curvature is controlled by the preset shape; further fine‑tuning would require
            // a newer API version that exposes adjustment values.

            // Save the workbook.
            workbook.Save("CurvedWordArt.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
