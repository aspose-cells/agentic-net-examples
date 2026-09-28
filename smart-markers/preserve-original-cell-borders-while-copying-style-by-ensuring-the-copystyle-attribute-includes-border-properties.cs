// Title: Copy a cell's style including border colors using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells Style.Copy to transfer all formatting, including border colors, from cell A1 to cell B2. | Provide a C# example that employs Aspose.Cells Range.CopyStyle to copy border styles from a source range to a destination range and then saves the file. | Write a C# snippet that applies the borders of a styled source cell to every cell in column C using Aspose.Cells style cloning methods.
// Common Searches: Aspose.Cells C# copy cell style while keeping border colors | Preserve border formatting when copying styles with Aspose.Cells | Example of Range.CopyStyle that retains borders in Aspose.Cells | How to duplicate full cell formatting, including borders, in Aspose.Cells workbook
// Tags: Aspose.Cells style copy with borders | C# Range.CopyStyle preserving borders | cell border formatting transfer Aspose.Cells | copy full cell formatting Aspose.Cells | save workbook with styled cells Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

// The example creates a workbook, styles cell A1 with distinct colored thin borders, copies the complete style—including those borders—to cell B2 using Style.Copy, optionally demonstrates copying via Range.CopyStyle, and saves the result as PreserveBorders.xlsx.
class PreserveBordersExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ---------- Source cell ----------
        // Set a value in the source cell
        Cell srcCell = cells["A1"];
        srcCell.PutValue("Source");

        // Create a style for the source cell and define borders
        Style srcStyle = workbook.CreateStyle();
        srcStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Red);
        srcStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Green);
        srcStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Blue);
        srcStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Orange);
        // Ensure border formatting is applied
        srcStyle.IsBorderApplied = true;

        // Apply the style to the source cell
        srcCell.SetStyle(srcStyle);

        // ---------- Destination cell ----------
        // Set a value in the destination cell
        Cell destCell = cells["B2"];
        destCell.PutValue("Destination");

        // Copy the entire style (including borders) from the source cell
        Style destStyle = workbook.CreateStyle();
        destStyle.Copy(srcCell.GetStyle());

        // Apply the copied style to the destination cell
        destCell.SetStyle(destStyle);

        // ---------- Alternative method using Range.CopyStyle ----------
        // Uncomment the following lines to copy style via ranges instead of individual cells
        // Range srcRange = cells.CreateRange("A1");
        // Range destRange = cells.CreateRange("C3");
        // destRange.CopyStyle(srcRange);

        // Save the workbook
        workbook.Save("PreserveBorders.xlsx");
    }
}
