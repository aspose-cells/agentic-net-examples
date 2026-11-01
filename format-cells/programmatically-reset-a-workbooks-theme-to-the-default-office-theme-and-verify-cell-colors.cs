// Title: Reset a workbook's theme to the default Office theme and verify cell colors with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# reset workbook theme to default Office theme | How to revert a cell's style to the workbook's default using Aspose.Cells | Check if cell font color returns to black after applying DefaultStyle in Aspose.Cells | Programmatically clear cell formatting to default theme with Aspose.Cells .NET | Validate default background pattern after resetting workbook theme in C#
// Tags: reset workbook theme Aspose.Cells | apply workbook.DefaultStyle C# | verify default cell font color Aspose.Cells | check cell background default Aspose.Cells | clear cell formatting to default theme

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a Workbook, changes cell A1's font to red and its fill to solid yellow, then restores the cell's style by assigning the workbook's DefaultStyle. It verifies that the font color reverts to black and the background pattern becomes none or transparent, outputting the verification results.
class ResetWorkbookTheme
{
    static void Main()
    {
        try
        {
            // Create a new workbook which uses the default Office theme
            Workbook workbook = new Workbook();

            // OPTIONAL: modify a cell to have a non‑default color to demonstrate the reset
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            Style style = cell.GetStyle();
            style.Font.Color = Color.Red;               // set font color to red
            style.ForegroundColor = Color.Yellow;       // set background color to yellow
            style.Pattern = BackgroundType.Solid;
            cell.SetStyle(style);

            // Reset the cell style to the workbook's default style (default theme)
            cell.SetStyle(workbook.DefaultStyle);

            // Verify that the cell colors have returned to the defaults (black font, no fill)
            Style resetStyle = cell.GetStyle();

            // Font color should be black (default)
            bool isFontDefault = resetStyle.Font.Color.ToArgb() == Color.Black.ToArgb();

            // Background should be no fill (default is Transparent)
            bool isBackgroundDefault = resetStyle.Pattern == BackgroundType.None ||
                                       resetStyle.ForegroundColor.ToArgb() == Color.Transparent.ToArgb();

            Console.WriteLine("Font color reset to default: " + isFontDefault);
            Console.WriteLine("Background reset to default: " + isBackgroundDefault);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
