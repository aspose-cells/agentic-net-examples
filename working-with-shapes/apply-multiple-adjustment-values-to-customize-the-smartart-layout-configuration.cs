// Title: Apply Multiple Geometry Adjustment Values to SmartArt Shapes with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, scans every worksheet for SmartArt shapes, accesses each shape's ShapeGuideCollection via shape.Geometry.ShapeAdjustValues, adds four custom adjustment entries (adj1‑adj4) with numeric values, and saves the file using OoxmlSaveOptions.UpdateSmartArt to persist the layout changes.
// Keywords: Aspose.Cells SmartArt adjustment | C# ShapeGuideCollection | UpdateSmartArt option | programmatic SmartArt geometry | Excel shape guides .NET | batch SmartArt customization | Aspose.Cells example
// Common Searches: how to set SmartArt adjustment values in Aspose.Cells C# | add multiple shape guides to Excel SmartArt using .NET | save workbook with updated SmartArt layout Aspose | modify SmartArt geometry programmatically | Aspose.Cells guide collection example
// Developer Intent: Programmatically assign several geometry adjustment parameters to SmartArt objects and write the modified workbook back to disk.
// Use Cases: Uniformly tweak roundness, angle, or other visual properties of all SmartArt diagrams in a report. | Automate layout fine‑tuning for dashboards that generate Excel files on the fly. | Integrate data‑driven adjustment values (e.g., from a database or API) into SmartArt before distribution.
// AI Prompts: Generate C# code that reads adjustment names and values from a JSON file and applies them to every SmartArt shape using Aspose.Cells. | Explain how to enumerate the available adjustment identifiers for a given SmartArt layout with Aspose.Cells. | Show how to export a SmartArt shape to PNG after applying adjustments to verify the visual result.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace AsposeCellsSmartArtAdjustDemo
{
    // Loads an Excel workbook, scans every worksheet for SmartArt shapes, accesses each shape's ShapeGuideCollection via shape.Geometry.ShapeAdjustValues, adds four custom adjustment entries (adj1‑adj4) with numeric values, and saves the file using OoxmlSaveOptions.UpdateSmartArt to persist the layout changes.
    class Program
    {
        static void Main()
        {
            const string inputPath = "SmartArtTemplate.xlsx";
            const string outputPath = "SmartArtAdjusted.xlsx";

            // Verify that the template file exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook that contains a SmartArt shape.
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet.
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only SmartArt shapes.
                        if (shape.IsSmartArt)
                        {
                            // Access the geometry adjustment collection of the SmartArt shape.
                            ShapeGuideCollection guides = shape.Geometry.ShapeAdjustValues;

                            // Example adjustments – names depend on the specific SmartArt layout.
                            guides.Add("adj1", 0.2); // First adjustment (e.g., roundness)
                            guides.Add("adj2", 0.4); // Second adjustment (e.g., angle)
                            guides.Add("adj3", 0.6); // Third adjustment
                            guides.Add("adj4", 0.8); // Fourth adjustment
                        }
                    }
                }

                // Save the workbook with SmartArt updates enabled.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime errors and display a friendly message.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
