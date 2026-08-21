// Title: Create a Union Range (B2:B10,F2:F10) and Apply Light Yellow Fill with Aspose.Cells for .NET
// Description: Shows how to use WorksheetCollection.CreateUnionRange in Aspose.Cells for .NET to merge the non‑contiguous ranges B2:B10 and F2:F10, apply a solid light‑yellow background style, and save the workbook.
// Keywords: Aspose.Cells | CreateUnionRange | union range C# | non‑adjacent cells | apply fill color | light yellow style | WorksheetCollection | Excel formatting .NET | solid background | C# Aspose.Cells example
// Common Searches: Aspose.Cells create union range C# | how to fill non‑contiguous cells with color Aspose.Cells | WorksheetCollection.CreateUnionRange usage | set background color for multiple ranges Aspose | C# Aspose.Cells light yellow fill
// Developer Intent: Combine the ranges B2:B10 and F2:F10 into a single union range and set a light‑yellow background using Aspose.Cells for .NET.
// Use Cases: Highlight two separate column sections in a financial report with a uniform color. | Apply consistent styling to non‑adjacent data columns that belong to the same category. | Create an input template where specific columns are colored to guide user entry.
// AI Prompts: Provide a C# snippet that uses WorksheetCollection.CreateUnionRange to merge B2:B10 and F2:F10 and apply a solid light‑yellow background with Aspose.Cells. | Show how to define a reusable style in Aspose.Cells for .NET and apply it to a union range of non‑contiguous cells. | Explain how to extend the light‑yellow style to other ranges after it has been applied to a union range in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Shows how to use WorksheetCollection.CreateUnionRange in Aspose.Cells for .NET to merge the non‑contiguous ranges B2:B10 and F2:F10, apply a solid light‑yellow background style, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a union range that combines B2:B10 and F2:F10 on the first sheet
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("B2:B10,F2:F10", 0);

        // Define a style with a solid light yellow fill
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightYellow;

        // Apply the style to the entire union range
        unionRange.ApplyStyle(style, new StyleFlag { All = true });

        // Save the workbook to a file
        workbook.Save("UnionRangeLightYellow.xlsx");
    }
}
