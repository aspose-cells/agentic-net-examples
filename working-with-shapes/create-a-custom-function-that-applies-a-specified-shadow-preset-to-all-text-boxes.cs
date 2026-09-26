// Title: How to apply a predefined shadow preset to every TextBox shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate a C# routine that traverses all worksheets in an Aspose.Cells workbook and assigns a top‑left drop‑shadow preset to each TextBox shape via the ShadowEffect object. | Create a reusable method that accepts a ShadowPreset enum value and updates the Blur, Distance, Angle, and Transparency properties of the ShadowEffect for all TextBox shapes in an Excel file using Aspose.Cells. | Extend the shadow‑application function to support additional presets such as inner shadow and perspective shadow, and show how to invoke it from the Main method.
// Common Searches: Aspose.Cells C# apply drop shadow to all text box shapes in a workbook | C# loop through worksheets and set ShadowEffect properties for TextBox using Aspose.Cells | How to batch update text box shadows in Excel with the Aspose.Cells ShadowEffect API | Define a ShadowPreset enum for shape formatting in Aspose.Cells C#
// Tags: apply shadowpreset to textbox shapes Aspose.Cells | iterate workbook shapes C# Aspose.Cells | set shadoweffect properties programmatically Excel | batch update shape shadows Aspose.Cells | shadowpreset enum for Aspose.Cells shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowExample
{
    // Simple enum to represent desired shadow presets.
    // Extend with additional presets as needed.
    internal enum ShadowPreset
    {
        TopLeftDropShadow
    }

    // The example loads an Excel workbook, iterates through each worksheet and its shapes, applies a predefined top‑left drop‑shadow preset to every TextBox via the ShadowEffect API, and saves the modified file.
    class Program
    {
        // Applies the given shadow preset to every text box in the workbook.
        static void ApplyShadowToAllTextBoxes(Workbook workbook, ShadowPreset preset)
        {
            try
            {
                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet.
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only TextBox shapes.
                        if (shape is TextBox)
                        {
                            // Access the shadow effect of the shape.
                            ShadowEffect shadow = shape.ShadowEffect;

                            // Apply preset‑specific shadow settings.
                            switch (preset)
                            {
                                case ShadowPreset.TopLeftDropShadow:
                                    shadow.Blur = 5;
                                    shadow.Distance = 5;
                                    shadow.Angle = 45; // Top‑left direction.
                                    shadow.Transparency = 0.5;
                                    break;

                                // Add more cases for other presets here.
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any errors that occur while applying shadows.
                Console.WriteLine($"Error applying shadows: {ex.Message}");
            }
        }

        static void Main()
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook.
                Workbook wb = new Workbook(inputPath);

                // Choose the desired shadow preset.
                ShadowPreset desiredPreset = ShadowPreset.TopLeftDropShadow;

                // Apply the shadow preset to all text boxes.
                ApplyShadowToAllTextBoxes(wb, desiredPreset);

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook.
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
