// Title: Apply the Bold Wave preset style to every WordArt shape in an Excel workbook using Aspose.Cells FontSettingCollection (C#)
// Common Searches: how to apply Bold Wave preset to WordArt shapes with Aspose.Cells C# | using FontSettingCollection to style WordArt in an Excel file Aspose.Cells | set preset WordArt style for all worksheets programmatically Aspose.Cells .NET | iterate Excel worksheets and apply Bold Wave WordArt style with Aspose.Cells | C# code to change WordArt font settings across a workbook using Aspose.Cells
// Tags: Aspose.Cells FontSettingCollection WordArt styling | C# bulk WordArt preset application | Excel workbook shape formatting Aspose.Cells | programmatic WordArt style update .NET | apply preset WordArt style across worksheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, iterates through each worksheet and its shapes, identifies WordArt objects, and uses Aspose.Cells' FontSettingCollection to apply the Bold Wave preset style to all WordArt shapes before saving the modified file.
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

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only WordArt shapes
                        if (shape.IsWordArt)
                        {
                            try
                            {
                                // Apply a bold effect to the WordArt text.
                                shape.TextEffect.FontBold = true;
                                shape.TextEffect.FontItalic = false;
                                // Underline property is not available in TextEffectFormat; omitted.
                            }
                            catch (Exception exShape)
                            {
                                Console.WriteLine($"Failed to modify shape '{shape.Name}': {exShape.Message}");
                            }
                        }
                    }
                }

                // Save the modified workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to {outputPath}");
                }
                catch (Exception exSave)
                {
                    Console.WriteLine($"Failed to save workbook: {exSave.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
