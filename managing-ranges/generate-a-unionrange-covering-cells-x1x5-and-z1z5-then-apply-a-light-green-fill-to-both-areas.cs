// Title: Aspose.Cells C# – Create UnionRange X1:X5, Z1:Z5 and Apply Light‑Green Fill
// Description: Shows how to build a UnionRange that merges the non‑adjacent cells X1:X5 and Z1:Z5 on the first worksheet, define a solid light‑green style, apply it to the range, and save the file as UnionRangeLightGreen.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | UnionRange | non‑adjacent range | light green fill | cell styling | Excel formatting | BackgroundType.Solid | StyleFlag | CreateUnionRange
// Common Searches: Aspose.Cells create union range | C# apply background color to multiple ranges | Style non‑contiguous cells Aspose.Cells | UnionRange light green example | How to set fill color for union range in .NET
// Developer Intent: Create a UnionRange for X1:X5 and Z1:Z5 and set a solid light‑green background on all cells.
// Use Cases: Highlight separate columns in a financial dashboard with a uniform color. | Design input‑output sections in a template where the same styling is required for non‑adjacent cells. | Prepare a worksheet for data entry by visually grouping related ranges.
// AI Prompts: Write C# code that creates a UnionRange covering X1:X5 and Z1:Z5 and applies a solid light‑green fill using Aspose.Cells. | Show how to style multiple non‑contiguous Excel ranges with a background color in Aspose.Cells for .NET. | Provide an example of defining a custom style and applying it to a UnionRange, then saving the workbook.

using System;
using Aspose.Cells;
using System.Drawing;

// Shows how to build a UnionRange that merges the non‑adjacent cells X1:X5 and Z1:Z5 on the first worksheet, define a solid light‑green style, apply it to the range, and save the file as UnionRangeLightGreen.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a union range that covers X1:X5 and Z1:Z5
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("X1:X5,Z1:Z5", 0);

        // Define a style with a solid light green fill
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightGreen;

        // Apply the style to the entire union range
        unionRange.ApplyStyle(style, new StyleFlag { All = true });

        // Save the workbook
        workbook.Save("UnionRangeLightGreen.xlsx");
    }
}
