// Title: C# – Set Uniform Shape Adjustment Guides for All Shapes Using Aspose.Cells
// Description: This example creates a workbook, adds several auto‑shapes, then iterates through every shape on the first worksheet. It accesses each shape’s Geometry.ShapeAdjustValues collection and assigns the same adjustment value (e.g., 0.3) to all existing guides, adding a default “adj” guide when a shape has none, before saving the file.
// Keywords: Aspose.Cells | C# | shape adjustment guide | uniform guide value | ShapeGuideCollection | modify shape geometry | Excel shape automation | loop through worksheet shapes
// Common Searches: Aspose.Cells set same adjustment guide for all shapes | C# loop worksheet shapes adjust geometry | add default shape guide Aspose.Cells | uniform shape adjustments Excel C# | change shape geometry programmatically Aspose.Cells
// Developer Intent: Apply a single adjustment value to every shape’s guides in a worksheet, creating a default guide when none exist.
// Use Cases: Standardize the look of callout or arrow shapes across a workbook. | Ensure custom geometry shapes have a baseline adjustment for consistent rendering. | Prepare Excel reports with uniform visual styling before distribution. | Automate bulk shape formatting in server‑side Excel generation.
// AI Prompts: Write C# code with Aspose.Cells that iterates over all worksheet shapes and sets each ShapeGuideCollection value to 0.3, adding an "adj" guide if the collection is empty. | Show how to retrieve and modify shape adjustment guides in Aspose.Cells, handling shapes without existing guides. | Explain the steps to apply a uniform adjustment value to every shape in an Excel file using Aspose.Cells and save the result.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// This example creates a workbook, adds several auto‑shapes, then iterates through every shape on the first worksheet. It accesses each shape’s Geometry.ShapeAdjustValues collection and assigns the same adjustment value (e.g., 0.3) to all existing guides, adding a default “adj” guide when a shape has none, before saving the file.
class ApplyUniformAdjustments
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample shapes that support adjustment guides
        worksheet.Shapes.AddAutoShape(AutoShapeType.RightArrowCallout, 2, 0, 2, 0, 200, 150);
        worksheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 5, 0, 5, 0, 200, 100);
        worksheet.Shapes.AddAutoShape(AutoShapeType.NotPrimitive, 8, 0, 8, 0, 200, 200); // custom geometry

        // Desired uniform adjustment value
        double uniformValue = 0.3;

        // Loop through all shapes in the worksheet
        for (int i = 0; i < worksheet.Shapes.Count; i++)
        {
            Shape shape = worksheet.Shapes[i];
            Geometry geometry = shape.Geometry;
            ShapeGuideCollection guides = geometry.ShapeAdjustValues;

            if (guides.Count > 0)
            {
                // Set each existing guide to the uniform value
                for (int j = 0; j < guides.Count; j++)
                {
                    guides[j].Value = uniformValue;
                }
            }
            else
            {
                // If no guides exist, add a default guide named "adj"
                guides.Add("adj", uniformValue);
            }
        }

        // Save the workbook with the updated shapes
        workbook.Save("UniformShapeAdjustments.xlsx");
    }
}
