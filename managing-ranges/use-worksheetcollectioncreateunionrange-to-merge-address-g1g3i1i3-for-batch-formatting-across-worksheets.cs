// Title: C# – Use WorksheetCollection.CreateUnionRange to batch‑format cells G1:G3 and I1:I3 in Aspose.Cells
// Description: Shows how to create a UnionRange for the non‑contiguous cells G1:G3 and I1:I3 on a worksheet, apply a solid light‑green background with bold font, assign a common value, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells UnionRange C# example | WorksheetCollection.CreateUnionRange | batch format non‑adjacent cells | apply style to union range | set value for union range | Aspose.Cells .NET styling | merge cell ranges programmatically | Excel union range Aspose
// Common Searches: Aspose.Cells create union range | How to style multiple non‑contiguous cells with one command | WorksheetCollection.CreateUnionRange C# tutorial | Set same value for several cells using Aspose.Cells | Batch formatting cells G1:G3 and I1:I3
// Developer Intent: Create a UnionRange that combines G1:G3 and I1:I3, apply a single style and optional value, then save the workbook.
// Use Cases: Standardize header colors across separate column groups in a worksheet. | Insert a common label into multiple cells of a financial report with one operation. | Apply background color and bold font to disjoint cells across many worksheets.
// AI Prompts: Generate C# code that creates a UnionRange for cells G1:G3 and I1:I3 on every worksheet and applies a red bold font style. | Show how to assign a formula to a UnionRange created with WorksheetCollection.CreateUnionRange and then save the workbook. | Provide an example that iterates through all worksheets, creates the same UnionRange, applies a style, and sets a text value.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsUnionRangeBatchFormatting
{
    // Shows how to create a UnionRange for the non‑contiguous cells G1:G3 and I1:I3 on a worksheet, apply a solid light‑green background with bold font, assign a common value, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a union range that combines G1:G3 and I1:I3 on the first worksheet
            // WorksheetCollection.CreateUnionRange(address, sheetIndex) returns a UnionRange object
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("G1:G3,I1:I3", 0);

            // Apply a style to the entire union range (e.g., light green background and bold font)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightGreen;
            style.Pattern = BackgroundType.Solid;
            style.Font.IsBold = true;
            // Apply the style to all formatting aspects
            unionRange.ApplyStyle(style, new StyleFlag { All = true });

            // Optionally set a value for the whole union range
            unionRange.Value = "Batch Formatted";

            // Save the workbook
            workbook.Save("UnionRangeBatchFormatting.xlsx");
        }
    }
}
