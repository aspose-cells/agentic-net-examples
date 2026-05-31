using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtBatchReplace
{
    class Program
    {
        static void Main()
        {
            const string templatePath = "TemplateWithSmartArt.xlsx";
            const string outputPath = "SmartArtReplaced.xlsx";

            // Verify that the template file exists before loading
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            try
            {
                // Load the workbook that contains SmartArt objects
                Workbook workbook = new Workbook(templatePath);

                // Text to find and its replacement
                const string placeholder = "OldPlaceholder";
                const string replacement = "NewValue";

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only SmartArt shapes
                        if (shape.IsSmartArt)
                        {
                            // Replace placeholder text directly in the shape's text
                            if (!string.IsNullOrEmpty(shape.Text) && shape.Text.Contains(placeholder))
                            {
                                shape.Text = shape.Text.Replace(placeholder, replacement);
                            }
                        }
                    }
                }

                // Save the workbook with SmartArt update enabled
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true // Ensure SmartArt text changes are persisted
                };
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}