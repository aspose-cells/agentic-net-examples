using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsMergePreserveStyle
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a 2x2 range with values and distinct styles
            // Cell A1
            Cell a1 = cells["A1"];
            a1.PutValue("Header");
            Style styleA1 = a1.GetStyle();
            styleA1.Font.IsBold = true;
            styleA1.Font.Color = Color.White;
            styleA1.ForegroundColor = Color.DarkBlue;
            styleA1.Pattern = BackgroundType.Solid;
            a1.SetStyle(styleA1);

            // Cell B1
            Cell b1 = cells["B1"];
            b1.PutValue("SubHeader");
            Style styleB1 = b1.GetStyle();
            styleB1.Font.IsItalic = true;
            styleB1.Font.Color = Color.Black;
            styleB1.ForegroundColor = Color.LightGray;
            styleB1.Pattern = BackgroundType.Solid;
            b1.SetStyle(styleB1);

            // Cell A2
            Cell a2 = cells["A2"];
            a2.PutValue(123);
            Style styleA2 = a2.GetStyle();
            styleA2.Font.Size = 12;
            styleA2.Font.Color = Color.Green;
            a2.SetStyle(styleA2);

            // Cell B2
            Cell b2 = cells["B2"];
            b2.PutValue(456);
            Style styleB2 = b2.GetStyle();
            styleB2.Font.Size = 12;
            styleB2.Font.Color = Color.Red;
            b2.SetStyle(styleB2);

            // Preserve the style of the top‑left cell (A1) before merging
            Style originalTopLeftStyle = a1.GetStyle();

            // Merge the 2x2 range (A1:B2) using Cells.Merge
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, checkConflict, mergeConflict
            // checkConflict = false (do not check for existing merges)
            // mergeConflict = true (allow merging even if conflicts exist)
            cells.Merge(0, 0, 2, 2, false, true);

            // After merging, reapply the original style to the merged cell.
            // Using SetStyle(style, true) ensures only explicitly set properties are overwritten,
            // leaving any default formatting untouched.
            Cell mergedCell = cells["A1"]; // Upper‑left cell of the merged area
            mergedCell.SetStyle(originalTopLeftStyle, true);

            // Save the workbook (lifecycle: save)
            workbook.Save("MergedPreserveStyle.xlsx", SaveFormat.Xlsx);
        }
    }
}