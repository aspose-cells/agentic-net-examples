// Title: Replace SmartArt node text in Excel with Aspose.Cells for .NET using a shape‑name dictionary
// Description: Loads an Excel workbook, builds a case‑insensitive Dictionary<string,string> that maps SmartArt shape names to new captions, walks every worksheet and each SmartArt shape, updates matching node texts, and saves the file with OoxmlSaveOptions.UpdateSmartArt so the changes persist.
// Keywords: Aspose.Cells | C# | .NET | SmartArt text replacement | UpdateSmartArt | shape name dictionary | Excel workbook automation | bulk SmartArt update | localize SmartArt labels
// Common Searches: change SmartArt node text Aspose.Cells C# | replace multiple SmartArt shapes using a dictionary | programmatically update SmartArt in Excel with .NET | Enable UpdateSmartArt when saving workbook | Aspose.Cells example for SmartArt text modification
// Developer Intent: Update specific SmartArt node captions based on a name‑to‑text mapping.
// Use Cases: Refresh diagram labels across all sheets in a financial report. | Localize SmartArt captions by swapping original names for translated strings. | Populate dashboard SmartArt titles from a data source via a dictionary lookup.
// AI Prompts: Show how to get the grouped shapes of a SmartArt object and replace their text with Aspose.Cells for .NET. | Create a reusable method that accepts a Workbook and a Dictionary<string,string> to update SmartArt node text and saves with UpdateSmartArt enabled. | Explain how to skip SmartArt shapes whose names are not present in the replacement dictionary while iterating.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace SmartArtTextReplacement
{
    // Loads an Excel workbook, builds a case‑insensitive Dictionary<string,string> that maps SmartArt shape names to new captions, walks every worksheet and each SmartArt shape, updates matching node texts, and saves the file with OoxmlSaveOptions.UpdateSmartArt so the changes persist.
    public class Program
    {
        public static void Main()
        {
            // Load the workbook (replace with your source file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Dictionary mapping SmartArt shape names to new text values
            var replacementMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "SmartArtNode1", "New Text 1" },
                { "SmartArtNode2", "New Text 2" },
                { "SmartArtNode3", "New Text 3" }
                // Add more mappings as needed
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Get the grouped shapes that represent the SmartArt nodes
                        Shape[] smartArtShapes = shape.GetResultOfSmartArt().GetGroupedShapes();

                        foreach (Shape smartArtShape in smartArtShapes)
                        {
                            // If the shape's name exists in the replacement dictionary, replace its text
                            if (replacementMap.TryGetValue(smartArtShape.Name, out string newText))
                            {
                                smartArtShape.Text = newText;
                            }
                        }
                    }
                }
            }

            // Save the workbook with SmartArt update enabled
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateSmartArt = true; // Ensure SmartArt text changes are persisted
            workbook.Save("output.xlsx", saveOptions);
        }
    }
}
