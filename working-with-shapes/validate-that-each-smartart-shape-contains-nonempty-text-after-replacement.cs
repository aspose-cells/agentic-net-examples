using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtValidation
{
    public class Validator
    {
        // Validates that every SmartArt shape in the workbook contains non‑empty text.
        // Throws an exception if any SmartArt sub‑shape has empty or whitespace text.
        public static void ValidateSmartArtText(string inputFilePath)
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputFilePath);

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Convert SmartArt to a group shape to access its constituent shapes
                        GroupShape group = shape.GetResultOfSmartArt();

                        // Guard against null (in case conversion fails)
                        if (group == null) continue;

                        // Examine each grouped shape that represents a SmartArt element
                        foreach (Shape smartArtPart in group.GetGroupedShapes())
                        {
                            // Check the Text property; consider null, empty or whitespace as invalid
                            if (string.IsNullOrWhiteSpace(smartArtPart.Text))
                            {
                                string message = $"SmartArt part with Id {smartArtPart.Id} in worksheet '{sheet.Name}' has empty text.";
                                throw new InvalidOperationException(message);
                            }
                        }
                    }
                }
            }

            // Save the workbook with UpdateSmartArt enabled (lifecycle rule: save)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateSmartArt = true; // ensure SmartArt changes are persisted
            workbook.Save(inputFilePath, saveOptions);
        }

        // Example usage
        public static void Main()
        {
            string filePath = "template.xlsx";

            try
            {
                ValidateSmartArtText(filePath);
                Console.WriteLine("All SmartArt shapes contain non‑empty text.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Validation failed: " + ex.Message);
            }
        }
    }
}