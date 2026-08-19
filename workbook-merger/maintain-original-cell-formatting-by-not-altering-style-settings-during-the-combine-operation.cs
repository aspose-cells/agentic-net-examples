// Title: Preserve Cell Formatting While Merging Ranges – Aspose.Cells C# Example
// Description: Shows how to keep the original style of the top‑left cell when merging a range (e.g., A1:B2) with Aspose.Cells for .NET. The sample captures the cell’s Style object, merges the cells, and reapplies the saved style so font, color, and background remain unchanged.
// Keywords: Aspose.Cells merge cells C# preserve style | keep formatting after merge Aspose.Cells | Aspose.Cells style retention | C# workbook merge preserve formatting | Aspose.Cells cell style object | merge range without losing formatting
// Common Searches: Aspose.Cells keep formatting after merge | C# merge cells preserve style Aspose | how to retain cell style when merging with Aspose.Cells | reapply style after merging cells .NET | preserve formatting of merged cells Aspose
// Developer Intent: Maintain the visual appearance of cells when combining them into a merged range.
// Use Cases: Create report headers that span multiple columns without losing bold text or background color. | Combine title and subtitle rows while preserving distinct font styles. | Generate printable spreadsheets where merged cells must match predefined branding.
// AI Prompts: Generate C# code that merges A1:B2 with Aspose.Cells and automatically retains the original style of A1. | Explain the steps to capture a cell’s Style, merge a range, and reapply the style using Aspose.Cells for .NET. | Provide a concise Aspose.Cells example that merges cells without altering font, color, or pattern.

using System.Drawing;
using Aspose.Cells;

// Shows how to keep the original style of the top‑left cell when merging a range (e.g., A1:B2) with Aspose.Cells for .NET. The sample captures the cell’s Style object, merges the cells, and reapplies the saved style so font, color, and background remain unchanged.
class PreserveFormattingMerge
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate cells with values and distinct styles
        cells["A1"].PutValue("Header");
        Style styleA1 = workbook.CreateStyle();
        styleA1.Font.IsBold = true;
        styleA1.Font.Color = Color.Blue;
        styleA1.ForegroundColor = Color.LightGray;
        styleA1.Pattern = BackgroundType.Solid;
        cells["A1"].SetStyle(styleA1);

        cells["B1"].PutValue("SubHeader");
        Style styleB1 = workbook.CreateStyle();
        styleB1.Font.IsItalic = true;
        styleB1.Font.Color = Color.Green;
        cells["B1"].SetStyle(styleB1);

        cells["A2"].PutValue("Data1");
        cells["B2"].PutValue("Data2");

        // Preserve the style of the top‑left cell before merging
        Style preservedStyle = cells["A1"].GetStyle();

        // Merge the range A1:B2 (rows 0‑1, columns 0‑1)
        cells.Merge(0, 0, 2, 2);

        // Reapply the preserved style to the merged cell (still addressed as A1)
        cells["A1"].SetStyle(preservedStyle);

        // Save the workbook
        workbook.Save("PreserveFormattingMerge.xlsx");
    }
}
