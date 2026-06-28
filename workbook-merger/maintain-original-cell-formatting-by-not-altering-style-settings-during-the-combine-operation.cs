using System.Drawing;
using Aspose.Cells;

class MaintainFormattingDuringMerge
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // -------------------------------------------------
        // Prepare sample cells with distinct formatting
        // -------------------------------------------------
        // Cell A1 – bold white text on dark blue background
        Cell cellA1 = cells["A1"];
        cellA1.PutValue("Header");
        Style styleA1 = cellA1.GetStyle();
        styleA1.Font.IsBold = true;
        styleA1.Font.Color = Color.White;
        styleA1.ForegroundColor = Color.DarkBlue;
        styleA1.Pattern = BackgroundType.Solid;
        cellA1.SetStyle(styleA1);

        // Cell B1 – italic black text on light gray background
        Cell cellB1 = cells["B1"];
        cellB1.PutValue("SubHeader");
        Style styleB1 = cellB1.GetStyle();
        styleB1.Font.IsItalic = true;
        styleB1.Font.Color = Color.Black;
        styleB1.ForegroundColor = Color.LightGray;
        styleB1.Pattern = BackgroundType.Solid;
        cellB1.SetStyle(styleB1);

        // -------------------------------------------------
        // Preserve original styles before merging
        // -------------------------------------------------
        Style savedStyleA1 = cellA1.GetStyle();
        // (savedStyleB1 is kept for reference; it will not be applied after merge
        // because merged cells collapse into a single cell)

        // -------------------------------------------------
        // Merge the range A1:B1 without altering formatting
        // -------------------------------------------------
        // Using Cells.Merge with mergeConflict = true to keep existing merged ranges intact
        cells.Merge(0, 0, 1, 2, true);

        // After merging, the merged cell is A1.
        // Reapply the original style of A1 using SetStyle with explicitFlag = true
        // This overwrites only the properties defined in savedStyleA1 and leaves others untouched.
        Cell mergedCell = cells["A1"];
        mergedCell.SetStyle(savedStyleA1, true);

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("MaintainFormattingMerge.xlsx", SaveFormat.Xlsx);
    }
}