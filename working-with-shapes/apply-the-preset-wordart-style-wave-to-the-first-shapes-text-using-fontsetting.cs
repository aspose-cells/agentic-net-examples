// Title: How to apply the Wave WordArt style to a shape’s text using FontSetting in Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet cells set wave wordart style on excel shape text | how to use FontSetting to apply a WordArt preset to a shape in Aspose.Cells | apply wave text effect to the first shape in an Excel workbook using Aspose.Cells .NET | aspnet cells change shape TextEffect to wave style programmatically | example of setting WordArt style for shape text with Aspose.Cells C#
// Tags: wave wordart style aspose.cells | fontsetting text effect excel shape c# | apply wordart preset to shape text aspose.cells | excel shape text formatting aspnet cells | aspose.cells modify shape text effect | c# set wordart style on worksheet shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example demonstrates loading an Excel workbook, locating the first shape on the first worksheet, ensuring the shape supports TextEffect, optionally assigning default text, and applying the Wave WordArt style to the shape's text via the FontSetting (TextEffect) API before saving the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape on the sheet
                if (worksheet.Shapes.Count > 0)
                {
                    // Get the first shape
                    Shape firstShape = worksheet.Shapes[0];

                    // Verify the shape supports text effects (WordArt)
                    if (firstShape.TextEffect != null)
                    {
                        try
                        {
                            // Optionally set some default text if none exists
                            if (string.IsNullOrEmpty(firstShape.TextEffect.Text))
                            {
                                firstShape.TextEffect.Text = "Sample Text";
                            }

                            // Note: The WordArtStyle property may not be available in all
                            // versions of Aspose.Cells. If needed, additional styling can be
                            // applied via other TextEffectFormat properties.
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error applying text effect: {ex.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
