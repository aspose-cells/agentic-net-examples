// Title: C# – Apply a Preset Shadow to All Text Boxes in an Aspose.Cells Worksheet
// Description: Demonstrates how to iterate through a worksheet's Shapes collection and set the ShadowEffect.PresetType for every text box using Aspose.Cells. The sample creates a workbook, adds three text boxes, applies the OffsetBottom preset (or any PresetShadowType), and saves the file as an XLSX document.
// Keywords: Aspose.Cells C# shadow effect | preset shadow text box | Worksheet.Shapes shadow preset | Apply ShadowEffect to shapes | PresetShadowType OffsetBottom | Excel text box styling | batch shape formatting Aspose
// Common Searches: how to set a preset shadow for all text boxes in Aspose.Cells C# | Aspose.Cells apply shadow effect to shapes worksheet | C# iterate worksheet shapes and change ShadowEffect | apply OffsetBottom shadow to multiple text boxes Excel | Aspose.Cells batch format text box shadows
// Developer Intent: Apply a uniform preset shadow to every text box in a worksheet programmatically.
// Use Cases: Standardize the appearance of callout boxes in auto‑generated reports. | Create a template where all annotation text boxes share the same shadow style for brand consistency. | Update existing Excel files in bulk to ensure every text box uses the same shadow preset.
// AI Prompts: Generate C# code that applies a chosen PresetShadowType only to text box shapes in an Aspose.Cells worksheet, ignoring other shape types. | Show how to call ApplyShadowToAllTextBoxes with different PresetShadowType values (e.g., OffsetBottom, OuterShadow) and export the workbook to PDF. | Explain how to modify ApplyShadowToAllTextBoxes to accept an array of shape IDs for selective shadow application.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to iterate through a worksheet's Shapes collection and set the ShadowEffect.PresetType for every text box using Aspose.Cells. The sample creates a workbook, adds three text boxes, applies the OffsetBottom preset (or any PresetShadowType), and saves the file as an XLSX document.
    public class TextBoxShadowHelper
    {
        // Applies the given preset shadow type to every shape (including text boxes) in the specified worksheet.
        public static void ApplyShadowToAllTextBoxes(Worksheet sheet, PresetShadowType preset)
        {
            foreach (Shape shape in sheet.Shapes)
            {
                // Apply the shadow preset to the shape.
                shape.ShadowEffect.PresetType = preset;
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a few text boxes for demonstration.
                sheet.Shapes.AddTextBox(1, 0, 1, 0, 150, 50);
                sheet.Shapes.AddTextBox(4, 0, 4, 0, 150, 50);
                sheet.Shapes.AddTextBox(7, 0, 7, 0, 150, 50);

                // Apply the desired shadow preset to all text boxes.
                ApplyShadowToAllTextBoxes(sheet, PresetShadowType.OffsetBottom);

                // Save the workbook.
                workbook.Save("TextBoxShadowDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            TextBoxShadowHelper.Run();
        }
    }
}
