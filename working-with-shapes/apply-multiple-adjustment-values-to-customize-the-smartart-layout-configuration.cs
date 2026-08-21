// Title: Set Multiple SmartArt Adjustment Values in Excel with Aspose.Cells for .NET
// Description: Loads or creates an Excel workbook, scans each worksheet for SmartArt shapes, converts them to GroupShape objects, adds custom adjustment guides (e.g., adj1, adj2, adj3) to the inner shapes' geometry, and saves the file with UpdateSmartArt enabled using Aspose.Cells.
// Keywords: Aspose.Cells | C# | SmartArt adjustment | ShapeAdjustValues | GroupShape | UpdateSmartArt | Excel automation | .NET | geometry guides | batch SmartArt processing
// Common Searches: Aspose.Cells set SmartArt adjustment values C# | How to modify SmartArt geometry with Aspose.Cells | Programmatically change SmartArt guides in Excel .NET | UpdateSmartArt option Aspose.Cells example | Apply multiple adjustment parameters to SmartArt shapes
// Developer Intent: Programmatically assign several adjustment parameters to every inner shape of a SmartArt diagram and persist the changes in the Excel file.
// Use Cases: Dynamically reshape SmartArt diagrams based on user input before distributing a template. | Standardize the appearance of SmartArt across a batch of workbooks by applying uniform adjustment guides. | Create automated reports where SmartArt proportions are tuned to reflect data-driven thresholds.
// AI Prompts: Write C# code with Aspose.Cells that iterates over all SmartArt shapes in a worksheet and sets custom adjustment values for each inner shape. | Show how to save an Excel workbook with the UpdateSmartArt flag after modifying SmartArt geometry. | Explain how to retrieve the list of adjustment guide names for a specific SmartArt layout using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace AsposeCellsSmartArtAdjustDemo
{
    // Loads or creates an Excel workbook, scans each worksheet for SmartArt shapes, converts them to GroupShape objects, adds custom adjustment guides (e.g., adj1, adj2, adj3) to the inner shapes' geometry, and saves the file with UpdateSmartArt enabled using Aspose.Cells.
    public class Program
    {
        public static void Main()
        {
            try
            {
                const string inputPath = "SmartArtTemplate.xlsx";
                const string outputPath = "SmartArtAdjusted.xlsx";

                // Ensure the input workbook exists; if not, create an empty workbook.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a new empty workbook.");
                    workbook = new Workbook();
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Convert the SmartArt to a group of shapes
                        GroupShape smartArtGroup = shape.GetResultOfSmartArt();

                        // Apply adjustment values to each inner shape
                        foreach (Shape innerShape in smartArtGroup.GetGroupedShapes())
                        {
                            Geometry geometry = innerShape.Geometry;

                            // Add adjustment guides (names depend on the specific SmartArt layout)
                            geometry.ShapeAdjustValues.Add("adj1", 0.2);
                            geometry.ShapeAdjustValues.Add("adj2", 0.5);
                            geometry.ShapeAdjustValues.Add("adj3", 0.8);
                        }
                    }
                }

                // Save the workbook with SmartArt updates enabled
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
