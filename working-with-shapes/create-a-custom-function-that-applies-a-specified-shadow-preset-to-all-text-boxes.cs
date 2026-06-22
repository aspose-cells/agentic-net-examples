using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    class Program
    {
        // Applies the given shadow preset to every text box in the workbook
        static void ApplyShadowPresetToAllTextBoxes(Workbook workbook, PresetShadowType preset)
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // If the shape contains text, treat it as a text box
                        if (!string.IsNullOrEmpty(shape.Text))
                        {
                            // Determine the length of the text
                            int textLength = shape.Text.Length;

                            // Get the FontSetting for the whole text range
                            FontSetting fontSetting = shape.Characters(0, textLength);

                            // Apply the shadow preset
                            fontSetting.TextOptions.Shadow.PresetType = preset;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log shape‑specific errors but continue processing other shapes
                        Console.WriteLine($"Warning: Could not apply shadow to a shape on sheet '{sheet.Name}'. Details: {ex.Message}");
                    }
                }
            }
        }

        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a few text boxes with sample text
                Shape tb1 = sheet.Shapes.AddTextBox(1, 1, 100, 30, 200, 100);
                tb1.Text = "First Box";

                Shape tb2 = sheet.Shapes.AddTextBox(5, 2, 150, 30, 200, 100);
                tb2.Text = "Second Box";

                // Apply a preset shadow (e.g., OffsetBottom) to all text boxes
                ApplyShadowPresetToAllTextBoxes(workbook, PresetShadowType.OffsetBottom);

                // Define output file path
                string outputPath = "AllTextBoxesWithShadow.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}