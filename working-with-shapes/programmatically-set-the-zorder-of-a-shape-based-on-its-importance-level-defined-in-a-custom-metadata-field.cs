// Title: How to set the Z‑order of Excel shapes from a custom ImportanceLevel stored in AlternativeText using Aspose.Cells for C#
// AI Prompts: Read each shape's AlternativeText, parse it as an integer, and assign that value to the shape's ZOrderPosition with Aspose.Cells in C#. | Iterate over all shapes on a worksheet and reorder them according to a numeric importance level defined in a custom metadata field. | Adapt the code to retrieve the importance level from a worksheet cell instead of AlternativeText and apply it to the shape's ZOrderPosition.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition based on custom metadata | Change Excel shape layering using AlternativeText property in Aspose.Cells | Programmatically reorder worksheet shapes by importance level with Aspose.Cells | How to bring important shapes to front using ZOrderPosition in Aspose.Cells | Read integer from shape.AlternativeText and set ZOrderPosition in C#
// Tags: Aspose.Cells shape ZOrderPosition | C# read shape AlternativeText | Excel shape layering based on metadata | set shape Z-order programmatically | custom importance level for Excel shapes

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The program loads an Excel workbook, iterates through every shape on the first worksheet, extracts an integer importance level from each shape's AlternativeText, assigns that value to the shape's ZOrderPosition to control layering, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Loop through all shapes on the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // Assume the custom metadata "ImportanceLevel" is stored in AlternativeText as an integer.
            // Example: shape.AlternativeText = "3";
            int importanceLevel = 0;
            if (int.TryParse(shape.AlternativeText, out importanceLevel))
            {
                // Set Z‑order based on importance.
                // In Aspose.Cells, ZOrderPosition = 0 is the backmost shape.
                // Higher values bring the shape to the front.
                shape.ZOrderPosition = importanceLevel;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
