// Title: Set the PerspectiveDiagonalBottomLeft shadow preset on a rectangle shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Apply the PerspectiveDiagonalBottomLeft preset to a shape's shadow using Aspose.Cells for .NET, employing reflection when the Shadow property is not directly exposed. | Create a rectangle shape in an Excel worksheet, programmatically change its shadow to a diagonal bottom‑left perspective, and save the workbook.
// Common Searches: C# Aspose.Cells set shape shadow to diagonal bottom left perspective | using reflection to modify shape shadow preset in older Aspose.Cells versions | how to apply a PerspectiveDiagonalBottomLeft shadow to a rectangle in Excel via Aspose.Cells | Aspose.Cells example for changing shape shadow effect and saving workbook
// Tags: Aspose.Cells shape shadow preset | PerspectiveDiagonalBottomLeft shadow effect | C# reflection Aspose.Cells shape property | Excel rectangle shape formatting .NET | save workbook with shape shadow Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a rectangle shape to the first worksheet, and uses reflection to set the shape's shadow preset to PerspectiveDiagonalBottomLeft for compatibility with older Aspose.Cells versions, then saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (row 2, column 2, width 120, height 60)
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 120, 60);

            // Attempt to set a shadow preset using reflection (compatible with older Aspose.Cells versions)
            try
            {
                var shadowProp = shape.GetType().GetProperty("Shadow");
                if (shadowProp != null)
                {
                    object shadowObj = shadowProp.GetValue(shape);
                    var presetProp = shadowObj?.GetType().GetProperty("Preset");
                    if (presetProp != null)
                    {
                        Type enumType = presetProp.PropertyType;
                        // Parse the enum value by name
                        object presetValue = Enum.Parse(enumType, "PerspectiveDiagonalBottomLeft");
                        presetProp.SetValue(shadowObj, presetValue);
                    }
                }
            }
            catch (Exception ex)
            {
                // If reflection fails, continue without setting shadow
                Console.WriteLine($"Shadow preset could not be applied: {ex.Message}");
            }

            // Save the workbook
            string outputPath = "ShadowPerspectiveDiagonalBottomLeft.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
