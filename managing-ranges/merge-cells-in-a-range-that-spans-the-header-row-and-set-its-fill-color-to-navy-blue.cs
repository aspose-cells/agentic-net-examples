// Title: Merge header row cells (A1:D1) and apply navy blue fill using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge cells in first row and set background color | How to apply solid navy fill to a merged header range with Aspose.Cells | Aspose.Cells style flag usage for merged cells in .NET | Create a merged header row and color it navy in a workbook using Aspose.Cells | Set foreground color to navy for merged cells in Aspose.Cells C# example
// Tags: merge header row Aspose.Cells C# | apply navy fill style Aspose.Cells | styleflag usage merged cells .NET | solid background color Aspose.Cells | range merge and style Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, merges cells A1:D1 on the first worksheet, defines a solid navy background style, applies the style to the merged header cell using a StyleFlag, and saves the file as MergedHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range that spans the header row (e.g., A1:D1)
        int startRow = 0;      // Row 1 (zero‑based)
        int startColumn = 0;   // Column A
        int endRow = 0;        // Row 1
        int endColumn = 3;     // Column D

        // Merge the cells in the defined range
        sheet.Cells.Merge(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);

        // Create a style and set its fill color to navy blue
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.Navy;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to the merged cell (top‑left cell of the range)
        StyleFlag flag = new StyleFlag();
        flag.All = true;
        sheet.Cells[startRow, startColumn].SetStyle(style, flag);

        // Save the workbook
        workbook.Save("MergedHeader.xlsx");
    }
}
