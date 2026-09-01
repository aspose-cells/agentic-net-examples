// Title: Set column C width to 20 points (≈0.28 inches) using Aspose.Cells in C#
// AI Prompts: Apply SetColumnWidthInch to column C (index 2) with a width of 20 points in a fresh Aspose.Cells workbook. | Convert 20 points to inches, set the column width, insert a long string into C1, and save the file as ColumnCWidth.xlsx. | Demonstrate column‑width adjustment by points in C# without referencing column letters.
// Common Searches: Aspose.Cells C# set column width in points | How to set Excel column C width to 20 points with Aspose.Cells | Set column width to 0.28 inches using SetColumnWidthInch method | Convert points to inches for column width in Aspose.Cells .NET | Adjust column width for long text entries in an Aspose.Cells workbook
// Tags: SetColumnWidthInch method usage | column width adjustment Aspose.Cells | points to inches conversion C# | Excel column sizing .NET API | long text column display Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, converts 20 points to inches, sets column C (index 2) width via SetColumnWidthInch, writes a long text into C1 to illustrate the width, and saves the workbook as ColumnCWidth.xlsx.
class SetColumnCWidth
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Column C is index 2 (0‑based). Set its width to 20 points.
        // 1 point = 1/72 inch, so 20 points = 20/72 inches.
        double inches = 20.0 / 72.0;
        worksheet.Cells.SetColumnWidthInch(2, inches);

        // Example data to visualize the width
        worksheet.Cells["C1"].PutValue("This is a long text entry that needs enough width.");

        // Save the workbook
        workbook.Save("ColumnCWidth.xlsx");
    }
}
