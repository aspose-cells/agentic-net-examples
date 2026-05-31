using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace AsposeCellsSmartArtAdjustDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains SmartArt shapes
            Workbook workbook = new Workbook("template.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Access the geometry of the SmartArt shape
                        Geometry geometry = shape.Geometry;

                        // Add multiple adjustment values (example values)
                        // The name can be "adj", "adj1", "adj2", etc., depending on the shape definition
                        geometry.ShapeAdjustValues.Add("adj1", 0.2);   // 20% adjustment
                        geometry.ShapeAdjustValues.Add("adj2", 0.5);   // 50% adjustment
                        geometry.ShapeAdjustValues.Add("adj3", 0.8);   // 80% adjustment

                        // Optionally, modify existing adjustment values
                        if (geometry.ShapeAdjustValues.Count > 0)
                        {
                            // Set the first adjustment to a new value
                            geometry.ShapeAdjustValues[0].Value = 0.35;
                        }
                    }
                }
            }

            // Save the workbook with SmartArt update enabled
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateSmartArt = true; // Apply the geometry changes to SmartArt

            workbook.Save("AdjustedSmartArt.xlsx", saveOptions);
        }
    }
}