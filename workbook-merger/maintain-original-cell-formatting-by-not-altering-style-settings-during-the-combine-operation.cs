// Title: How to merge a cell range in Aspose.Cells for .NET while keeping the original formatting of the top‑left cell
// AI Prompts: Generate C# code that merges cells A1:B2 with Aspose.Cells and restores the original style of the merged cell. | Show how to retrieve a cell's Style object before calling Cells.Merge and then reapply it after the merge in a .NET workbook. | Provide a step‑by‑step example of merging a range and preserving font, color, and background settings of the first cell using Aspose.Cells.
// Common Searches: Aspose.Cells C# merge cells without losing original cell style | retain font and background after merging range Aspose.Cells .NET | how to keep header formatting when merging A1:B2 in Aspose.Cells | reapply captured Style to merged cell using Aspose.Cells API | preserve cell formatting during merge operation in C# workbook
// Tags: Aspose.Cells merge preserve style | capture cell style before merge C# | reapply style after cell merge Aspose.Cells | merge range without formatting loss .NET | maintain original formatting merged cells

using Aspose.Cells;
using System.Drawing;

// Demonstrates merging the range A1:B2 in a new workbook with Aspose.Cells for .NET, capturing the Style of the top‑left cell before the merge, and reapplying it afterward so the original font, color, and background remain unchanged.
class PreserveFormattingDuringMerge
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with values and distinct styles
        cells["A1"].PutValue("Header");
        Style styleA1 = workbook.CreateStyle();
        styleA1.Font.IsBold = true;
        styleA1.Font.Color = Color.Blue;
        styleA1.ForegroundColor = Color.LightGray;
        styleA1.Pattern = BackgroundType.Solid;
        cells["A1"].SetStyle(styleA1);

        cells["A2"].PutValue("Data1");
        Style styleA2 = workbook.CreateStyle();
        styleA2.Font.Color = Color.Green;
        cells["A2"].SetStyle(styleA2);

        cells["B1"].PutValue("Header2");
        Style styleB1 = workbook.CreateStyle();
        styleB1.Font.IsItalic = true;
        cells["B1"].SetStyle(styleB1);

        cells["B2"].PutValue("Data2");
        // B2 keeps default style

        // Capture the original style of the top‑left cell (A1) before merging
        Style originalStyle = cells["A1"].GetStyle();

        // Merge the range A1:B2 (2 rows x 2 columns)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, checkConflict, mergeConflict
        cells.Merge(0, 0, 2, 2, true, true);

        // Reapply the captured style to the merged cell to keep formatting unchanged
        cells["A1"].SetStyle(originalStyle, true);

        // Save the workbook
        workbook.Save("PreserveFormattingMerge.xlsx");
    }
}
