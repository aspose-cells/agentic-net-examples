using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the workbook containing SmartArt
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary: key = SmartArt shape name, value = replacement text
        var replacements = new Dictionary<string, string>
        {
            { "SmartArtNode1", "New Text 1" },
            { "SmartArtNode2", "New Text 2" },
            // add more mappings as needed
        };

        // Iterate through all worksheets and shapes
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    // Get the individual shapes that make up the SmartArt
                    foreach (Shape smartArtShape in shape.GetResultOfSmartArt().GetGroupedShapes())
                    {
                        // If the shape name exists in the dictionary, replace its text
                        if (replacements.TryGetValue(smartArtShape.Name, out string newText))
                        {
                            smartArtShape.Text = newText;
                        }
                    }
                }
            }
        }

        // Save the workbook with SmartArt updates applied
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.UpdateSmartArt = true;
        workbook.Save("output.xlsx", saveOptions);
    }
}