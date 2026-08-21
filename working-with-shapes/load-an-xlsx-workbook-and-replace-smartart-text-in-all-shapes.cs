// Title: C# – Replace SmartArt and Alternative Text in All Shapes of an XLSX Workbook with Aspose.Cells
// Description: Loads an XLSX file, iterates every worksheet and each shape, sets a new AlternativeText for all shapes, detects SmartArt objects, converts them to GroupShape, replaces the text of every grouped shape originating from SmartArt, and saves the workbook with the UpdateSmartArt option enabled.
// Keywords: Aspose.Cells C# | SmartArt text replacement | Excel shape alternative text | Iterate worksheet shapes | UpdateSmartArt option | GroupShape Aspose | Bulk shape text update | C# Excel automation
// Common Searches: Aspose.Cells replace SmartArt text C# | How to change alternative text for all shapes in Excel using Aspose | Save workbook with UpdateSmartArt flag | Convert SmartArt to GroupShape Aspose.Cells | Iterate shapes on each worksheet C#
// Developer Intent: Update every shape’s alternative text and replace all SmartArt node text in an XLSX file, then save with SmartArt refreshed.
// Use Cases: Refresh SmartArt labels in a reporting template before distribution. | Add descriptive alternative text to all diagram elements for accessibility compliance. | Automate bulk editing of shape captions after data‑driven calculations. | Create a localized version of a workbook by swapping SmartArt text in one step.
// AI Prompts: Generate code that reads a CSV mapping and replaces SmartArt node text accordingly. | Show how to log shape name, type, and original text before modification. | Explain error‑handling for worksheets without shapes or SmartArt. | Provide a version that processes only selected worksheets based on their name.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an XLSX file, iterates every worksheet and each shape, sets a new AlternativeText for all shapes, detects SmartArt objects, converts them to GroupShape, replaces the text of every grouped shape originating from SmartArt, and saves the workbook with the UpdateSmartArt option enabled.
class Program
{
    static void Main()
    {
        // Load the source workbook (lifecycle rule: load)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Example: replace alternative text of every shape
                shape.AlternativeText = "ReplacedAlternativeText";

                // Check if the shape is a SmartArt object
                if (shape.IsSmartArt)
                {
                    // Convert SmartArt to a grouped shape (cached result)
                    GroupShape groupShape = shape.GetResultOfSmartArt();

                    // Replace text in each grouped shape that originated from SmartArt
                    foreach (Shape smartArtShape in groupShape.GetGroupedShapes())
                    {
                        smartArtShape.Text = "ReplacedSmartArtText";
                    }
                }
            }
        }

        // Save the workbook with SmartArt update enabled (lifecycle rule: save)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.UpdateSmartArt = true;
        workbook.Save("output.xlsx", saveOptions);
    }
}
