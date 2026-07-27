// Title: Disable Shape Reflections in Excel with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, iterates through every worksheet and shape, sets each shape's Reflection.Type to None, and saves the workbook, leaving all other visual effects untouched.
// Keywords: Aspose.Cells C# disable reflection | remove shape reflection Excel | preserve shape effects Aspose | iterate shapes worksheet | ReflectionEffectType.None example
// Common Searches: how to turn off reflections on all Excel shapes using Aspose.Cells | C# code to remove shape reflection while keeping other effects | Aspose.Cells loop through shapes and disable reflection
// Developer Intent: Programmatically eliminate reflection effects from every shape in a workbook without altering other formatting.
// Use Cases: Prepare a report for printing where reflections cause visual noise. | Standardize the appearance of templates before distribution. | Automate cleanup of legacy workbooks that contain unwanted shape reflections.
// AI Prompts: Generate C# code that disables reflections on all shapes in an Excel workbook using Aspose.Cells and preserves other effects. | Explain the steps to loop through worksheets and shapes to set Reflection.Type to None in Aspose.Cells. | Show how to verify a shape has a reflection before clearing it to avoid redundant assignments.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file, iterates through every worksheet and shape, sets each shape's Reflection.Type to None, and saves the workbook, leaving all other visual effects untouched.
class DisableShapeReflections
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through all shapes in the current worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Get the reflection effect of the shape
                ReflectionEffect reflection = shape.Reflection;

                // Disable the reflection while keeping other effects intact
                reflection.Type = ReflectionEffectType.None;
            }
        }

        // Save the workbook with reflections disabled
        workbook.Save("output.xlsx");
    }
}
