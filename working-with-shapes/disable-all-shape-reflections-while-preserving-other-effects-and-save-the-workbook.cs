// Title: C# – Disable All Shape Reflections in an Excel Workbook with Aspose.Cells (Preserve Other Effects)
// Description: Loads a workbook, walks through every worksheet and shape, sets each shape's Reflection.Type to None to turn off reflections while keeping all other visual effects, and saves the result to a new file.
// Keywords: Aspose.Cells shape reflection | disable shape reflection .NET | Reflection.Type None | iterate shapes Aspose.Cells | preserve shape effects Excel
// Common Searches: how to remove reflections from all shapes using Aspose.Cells | Aspose.Cells disable shape reflection C# | turn off shape reflections while keeping other effects | batch remove Excel shape reflections Aspose
// Developer Intent: Remove the reflection effect from every shape in an Excel workbook without altering any other formatting or visual properties.
// Use Cases: Prepare a printable report where shape reflections cause visual noise. | Standardize the appearance of templates before distribution by disabling reflections. | Automate batch processing of multiple workbooks to ensure consistent shape styling.
// AI Prompts: Write C# code with Aspose.Cells that disables reflections on all shapes in a workbook and saves the file. | Explain how to modify only the reflection property of shapes while leaving other effects untouched in Aspose.Cells for .NET. | Extend the sample to log each shape's name or ID when its reflection is turned off.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, walks through every worksheet and shape, sets each shape's Reflection.Type to None to turn off reflections while keeping all other visual effects, and saves the result to a new file.
class DisableShapeReflections
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Access the reflection effect of the shape
                ReflectionEffect reflection = shape.Reflection;

                // Disable reflection by setting its type to None
                reflection.Type = ReflectionEffectType.None;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
