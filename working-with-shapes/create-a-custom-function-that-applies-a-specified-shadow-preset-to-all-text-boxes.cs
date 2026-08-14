// Title: Apply a Preset Shadow to Every Text Box in an Aspose.Cells Worksheet (C#)
// Description: Provides a reusable C# method that scans all shapes on a worksheet, detects text boxes via the TextBody property, and assigns a chosen PresetShadowType to each shape's ShadowEffect. The example adds two text boxes and a rectangle, applies the OffsetBottom preset to the text boxes only, and saves the workbook.
// Keywords: Aspose.Cells | C# | text box shadow | PresetShadowType | ShadowEffect | apply shadow to shapes | iterate worksheet shapes | Excel workbook styling | OffsetBottom preset | custom shape helper
// Common Searches: Aspose.Cells set preset shadow for all text boxes C# | how to apply shadow effect only to text boxes in Excel using Aspose.Cells | C# iterate worksheet shapes and change ShadowEffect | apply OffsetBottom shadow to text boxes Aspose.Cells | sample code for text box shadow preset Aspose.Cells
// Developer Intent: Assign a specific preset shadow to every text box on a worksheet while preserving other shape formats.
// Use Cases: Standardize call‑out box appearance in automated reports by applying a uniform shadow preset. | Create a template where newly added text boxes inherit a predefined shadow without affecting rectangles or charts. | Batch‑process workbooks to visually separate text boxes from other shapes through consistent shadow styling.
// AI Prompts: Generate C# code that defines a method to apply any PresetShadowType to all text boxes in an Aspose.Cells worksheet, including null checks and error handling. | Show how to invoke ApplyShadowPresetToAllTextBoxes with a user‑selected shadow type and log the names of modified shapes. | Explain how to extend the helper to skip text boxes that already have a shadow preset or to assign different presets based on the box content.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Provides a reusable C# method that scans all shapes on a worksheet, detects text boxes via the TextBody property, and assigns a chosen PresetShadowType to each shape's ShadowEffect. The example adds two text boxes and a rectangle, applies the OffsetBottom preset to the text boxes only, and saves the workbook.
    public class TextBoxShadowHelper
    {
        // Applies the given preset shadow type to every text box on the specified worksheet.
        public static void ApplyShadowPresetToAllTextBoxes(Worksheet sheet, PresetShadowType preset)
        {
            // Iterate through all shapes in the worksheet.
            foreach (Shape shape in sheet.Shapes)
            {
                // Text boxes have a TextBody (FontSettingCollection). If it exists, treat the shape as a text box.
                if (shape.TextBody != null)
                {
                    // Set the shadow preset for the shape.
                    shape.ShadowEffect.PresetType = preset;
                }
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a few sample text boxes.
                Shape tb1 = sheet.Shapes.AddTextBox(1, 1, 100, 100, 200, 50);
                tb1.TextBody.Text = "First Box";

                Shape tb2 = sheet.Shapes.AddTextBox(5, 2, 300, 150, 200, 50);
                tb2.TextBody.Text = "Second Box";

                // Add a non‑text‑box shape for contrast.
                Shape rect = sheet.Shapes.AddRectangle(8, 1, 500, 200, 100, 60);
                rect.TextBody.Text = "Rectangle";

                // Apply the desired shadow preset to all text boxes.
                ApplyShadowPresetToAllTextBoxes(sheet, PresetShadowType.OffsetBottom);

                // Save the workbook.
                workbook.Save("TextBoxesWithShadow.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully as TextBoxesWithShadow.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            TextBoxShadowHelper.Run();
        }
    }
}
