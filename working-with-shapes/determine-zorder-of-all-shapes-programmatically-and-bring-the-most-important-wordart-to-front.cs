// Title: Programmatically Determine Shape Z‑Order and Bring Top WordArt to Front with Aspose.Cells for .NET (C#)
// Description: This C# example loads an Excel workbook, scans every worksheet for shapes whose name contains "WordArt", identifies the shape with the highest ZOrderPosition, sets its ZOrderPosition to 0 to place it at the front, and saves the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel shape Z-order | WordArt front layer | shape layering .NET | ZOrderPosition | move shape to front | programmatic shape ordering
// Common Searches: Aspose.Cells set shape Z-order C# | bring WordArt to front Excel programmatically | find highest ZOrderPosition shape Aspose.Cells | change shape layering in .NET Excel | move specific shape to front using Aspose
// Developer Intent: Locate the WordArt shape with the greatest Z-order on each worksheet and reposition it as the frontmost object.
// Use Cases: Guarantee that a title WordArt appears above all other graphics when a report is opened. | Automate shape re‑ordering in generated workbooks so key annotations remain visible. | Prepare existing Excel files for publishing by ensuring priority WordArt is not obscured.
// AI Prompts: Create a reusable C# method that accepts a shape name pattern and moves the matching shape with the highest ZOrderPosition to the front using Aspose.Cells. | Generate code to list all shapes on each worksheet together with their ZOrderPosition values. | Provide an enhanced version of the sample that includes null‑checks, exception handling, and saves the workbook to a MemoryStream.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example loads an Excel workbook, scans every worksheet for shapes whose name contains "WordArt", identifies the shape with the highest ZOrderPosition, sets its ZOrderPosition to 0 to place it at the front, and saves the updated file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Shape mostImportantWordArt = null;
            int highestZOrder = int.MinValue;

            // Examine all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Identify WordArt shapes (commonly named with "WordArt")
                if (shape.Name != null && shape.Name.IndexOf("WordArt", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Track the shape with the highest Z-order position
                    if (shape.ZOrderPosition > highestZOrder)
                    {
                        highestZOrder = shape.ZOrderPosition;
                        mostImportantWordArt = shape;
                    }
                }
            }

            // Bring the most important WordArt to the front
            if (mostImportantWordArt != null)
            {
                // Frontmost position is 0 (lower value means closer to front)
                mostImportantWordArt.ZOrderPosition = 0;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
