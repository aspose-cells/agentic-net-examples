// Title: Enable text wrap for a specific cell range (A1:C3) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# enable wrap text for a range of cells | How to apply text wrapping to multiple cells with Aspose.Cells API | Using StyleFlag to set IsTextWrapped on a cell block in Aspose.Cells | CreateRange ApplyStyle example for text wrap in Aspose.Cells .NET | Set wrap text style for A1:C3 in an Excel workbook with Aspose.Cells
// Tags: apply text wrap Aspose.Cells | StyleFlag ApplyStyle range | CreateRange text wrapping | IsTextWrapped property C# | wrap text Excel Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, defines the A1:C3 range, builds a style with IsTextWrapped enabled, uses a StyleFlag with All set to true, applies the style to the range via CreateRange.ApplyStyle, and saves the file as WrappedText.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell range (A1:C3) to apply text wrap
        int startRow = 0;      // Row 1 (zero‑based)
        int startColumn = 0;   // Column A
        int totalRows = 3;     // Rows 1‑3
        int totalColumns = 3;  // Columns A‑C

        // Create a style with text wrap enabled
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;

        // Define which style attributes to apply (all in this case)
        StyleFlag flag = new StyleFlag() { All = true };

        // Apply the wrap style to the specified range
        sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns).ApplyStyle(wrapStyle, flag);

        // Save the workbook
        workbook.Save("WrappedText.xlsx");
    }
}
