// Title: Set the First SmartArt Adjustment in Excel with Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans each worksheet for SmartArt shapes, accesses the Geometry.ShapeAdjustValues collection, changes the first adjustment value, and saves the file using OoxmlSaveOptions.UpdateSmartArt to persist the modification.
// Keywords: Aspose.Cells | C# | SmartArt adjustment | Shape.Adjustments | Geometry.ShapeAdjustValues | UpdateSmartArt | Excel automation | batch SmartArt editing | guide value | Excel shape programming
// Common Searches: how to modify SmartArt adjustment value with Aspose.Cells | Aspose.Cells change first guide of SmartArt diagram | C# update SmartArt geometry adjustments before saving | set SmartArt shape adjustment in Excel using .NET | Aspose.Cells UpdateSmartArt option example
// Developer Intent: Programmatically change the first adjustment (guide) of a SmartArt shape in an Excel workbook and save the updated file.
// Use Cases: Standardize SmartArt proportions across multiple reports by adjusting the primary guide value. | Create a template‑driven workflow that customizes SmartArt layouts before exporting to PDF. | Batch‑process a folder of workbooks to enforce a consistent SmartArt appearance for corporate branding.
// AI Prompts: Generate C# code that sets the second SmartArt adjustment to 0.75 and saves the workbook with UpdateSmartArt enabled. | Explain the purpose of Shape.Geometry.ShapeAdjustValues and show how to iterate through all adjustments of a SmartArt shape. | Add comprehensive error handling for cases where a SmartArt shape has no adjustment values or the workbook lacks SmartArt objects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, scans each worksheet for SmartArt shapes, accesses the Geometry.ShapeAdjustValues collection, changes the first adjustment value, and saves the file using OoxmlSaveOptions.UpdateSmartArt to persist the modification.
    public class ModifySmartArtAdjustment
    {
        public static void Run()
        {
            const string inputPath = "SmartArtTemplate.xlsx";
            const string outputPath = "ModifiedSmartArt.xlsx";

            try
            {
                // Verify that the input file exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load a workbook that contains a SmartArt shape.
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and shapes.
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        // Check if the shape is a SmartArt.
                        if (shape.IsSmartArt)
                        {
                            // Access the geometry adjustments collection.
                            ShapeGuideCollection adjustments = shape.Geometry.ShapeAdjustValues;

                            // Ensure there is at least one adjustment value.
                            if (adjustments.Count > 0)
                            {
                                // Modify the first adjustment value.
                                adjustments[0].Value = 0.5; // Set desired value.
                            }
                        }
                    }
                }

                // Save the workbook with SmartArt updates enabled.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ModifySmartArtAdjustment.Run();
        }
    }
}
