using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Lookup table: placeholder -> replacement text
        var lookup = new Dictionary<string, string>
        {
            { "{Name}", "John Doe" },
            { "{Date}", "2023-12-31" },
            { "{Amount}", "$123.45" }
        };

        // Iterate through all shapes in the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // Ensure the shape has a text body and some text
            if (shape.TextBody != null && !string.IsNullOrEmpty(shape.Text))
            {
                // Perform replacements based on the lookup table
                foreach (var kvp in lookup)
                {
                    if (shape.Text.Contains(kvp.Key))
                    {
                        // Replace the placeholder with the actual value
                        shape.TextBody.Replace(kvp.Key, kvp.Value);
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}