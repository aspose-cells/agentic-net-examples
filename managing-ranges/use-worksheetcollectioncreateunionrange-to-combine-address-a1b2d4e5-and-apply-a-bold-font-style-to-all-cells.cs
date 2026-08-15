// Title: Aspose.Cells C# – Create Union Range (A1:B2, D4:E5) and Apply Bold Font
// Description: Learn how to use WorksheetCollection.CreateUnionRange to combine the non‑contiguous cells A1:B2 and D4:E5, define a bold Style, apply it with a StyleFlag, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells CreateUnionRange | C# union range bold style | WorksheetCollection.CreateUnionRange | Apply bold font Aspose.Cells | StyleFlag FontBold | non contiguous cells formatting | Aspose.Cells .NET example | UnionRange.ApplyStyle
// Common Searches: How to create a union range with Aspose.Cells C# | Apply bold formatting to multiple non‑adjacent cells using Aspose.Cells | WorksheetCollection CreateUnionRange example | StyleFlag usage for font bold in Aspose.Cells
// Developer Intent: Combine A1:B2 and D4:E5 into a union range and set the font to bold for every cell in that range.
// Use Cases: Highlight header blocks that are separated on a worksheet by applying bold formatting to a union range. | Emphasize key financial figures located in different sections of a report without merging cells. | Standardize bold styling across several chart data ranges that are not contiguous.
// AI Prompts: Generate C# code that uses WorksheetCollection.CreateUnionRange to merge A1:B2 and D4:E5 and applies a bold font with a StyleFlag. | Show an Aspose.Cells example that applies only the FontBold attribute to a union range. | Explain how to reuse a StyleFlag to apply additional style properties to a union range created with CreateUnionRange.

using System;
using Aspose.Cells;

// Learn how to use WorksheetCollection.CreateUnionRange to combine the non‑contiguous cells A1:B2 and D4:E5, define a bold Style, apply it with a StyleFlag, and save the workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Create a union range that combines cells A1:B2 and D4:E5 on the first worksheet (index 0)
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,D4:E5", 0);

        // Define a style with a bold font
        Style boldStyle = workbook.CreateStyle();
        boldStyle.Font.IsBold = true;

        // Specify that only the bold attribute should be applied
        StyleFlag flag = new StyleFlag();
        flag.FontBold = true;

        // Apply the bold style to all cells in the union range
        unionRange.ApplyStyle(boldStyle, flag);

        // Save the workbook to a file
        workbook.Save("UnionRangeBold.xlsx");
    }
}
