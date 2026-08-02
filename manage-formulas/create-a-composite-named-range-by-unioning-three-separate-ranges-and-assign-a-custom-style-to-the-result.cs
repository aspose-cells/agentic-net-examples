// Title: C# – Create a Composite (Union) Named Range and Apply a Custom Style with Aspose.Cells for .NET
// Description: This example shows how to build a UnionRange that combines three separate areas (A1:B2, C3:D4, E5:F6), assign it the name "MyCompositeRange", define a solid light‑green, bold dark‑blue style with centered alignment, apply the style to the whole range, and save the workbook as CompositeNamedRange.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells UnionRange C# | composite named range .NET | noncontiguous cells style | apply custom style Aspose.Cells | create named range multiple areas | C# Excel formatting example | Aspose.Cells sample code | union range styling
// Common Searches: Aspose.Cells create named range from non‑adjacent cells | C# apply style to UnionRange in Excel | how to union multiple ranges in Aspose.Cells | custom cell formatting for composite range .NET | sample code for UnionRange with style
// Developer Intent: Generate a UnionRange that spans several non‑adjacent blocks, give it a name, and format the entire range with a single custom style.
// Use Cases: Define a report section that pulls data from scattered cells and formats them uniformly. | Set up a printable area composed of multiple separate blocks with consistent background and font. | Create a dashboard where related metrics reside in different ranges but share the same visual appearance.
// AI Prompts: Provide C# code that creates a UnionRange for "A1:B2,C3:D4,E5:F6", names it, and applies a solid light‑green, bold dark‑blue style using Aspose.Cells. | Show how to modify the style of an existing composite named range in Aspose.Cells for .NET. | Explain step‑by‑step how to combine non‑contiguous cells into a single named range and format them with one style in C#.

using System;
using System.Drawing;
using Aspose.Cells;

// This example shows how to build a UnionRange that combines three separate areas (A1:B2, C3:D4, E5:F6), assign it the name "MyCompositeRange", define a solid light‑green, bold dark‑blue style with centered alignment, apply the style to the whole range, and save the workbook as CompositeNamedRange.xlsx using Aspose.Cells for .NET.
class CompositeNamedRangeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in the three individual ranges
        worksheet.Cells["A1"].PutValue("Range1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["C3"].PutValue("Range2");
        worksheet.Cells["D4"].PutValue(200);
        worksheet.Cells["E5"].PutValue("Range3");
        worksheet.Cells["F6"].PutValue(300);

        // Create a composite (union) range that combines the three areas
        // Address format: "A1:B2,C3:D4,E5:F6"
        UnionRange compositeRange = workbook.Worksheets.CreateUnionRange("A1:B2,C3:D4,E5:F6", 0);

        // Assign a name to the composite range (named range)
        compositeRange.Name = "MyCompositeRange";

        // Define a custom style to apply to the whole union range
        Style customStyle = workbook.CreateStyle();
        customStyle.Pattern = BackgroundType.Solid;
        customStyle.ForegroundColor = Color.LightGreen;
        customStyle.Font.IsBold = true;
        customStyle.Font.Color = Color.DarkBlue;
        customStyle.HorizontalAlignment = TextAlignmentType.Center;
        customStyle.VerticalAlignment = TextAlignmentType.Center;

        // Apply the style to the composite range
        compositeRange.SetStyle(customStyle);

        // Save the workbook
        workbook.Save("CompositeNamedRange.xlsx");
    }
}
