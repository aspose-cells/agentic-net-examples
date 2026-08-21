// Title: Copy Cell Style with Borders Using Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, applies thick colored borders to cell A1, copies the complete style—including all border properties—to cell B2 using Style.Copy (or Range.CopyStyle), verifies the border colors, and saves the file as StyleCopyWithBorders.xlsx.
// Keywords: Aspose.Cells | .NET | C# | copy cell style | preserve borders | Style.Copy | Range.CopyStyle | Excel formatting | border colors | cell style duplication
// Common Searches: Aspose.Cells copy style with borders C# | How to preserve cell borders when copying style Aspose.Cells | Style.Copy includes border properties | Range.CopyStyle border preservation | Copy cell formatting Aspose.Cells .NET | Copy borders from one cell to another using Aspose.Cells
// Developer Intent: Copy a cell’s formatting, including all border attributes, to another cell in an Excel workbook.
// Use Cases: Create a template cell with custom thick colored borders and reuse its style across multiple report cells. | Duplicate header or footer cell border formatting programmatically when generating Excel invoices. | Apply consistent border styling to a dynamic data range using Range.CopyStyle in bulk operations. | Adjust border colors after copying a style for conditional formatting scenarios.
// AI Prompts: Generate C# code that copies a cell’s style with all border settings using Aspose.Cells Style.Copy and prints the border colors. | Show an example of copying borders from one range to another with Aspose.Cells Range.CopyStyle in .NET. | Explain how to modify border colors after a style has been copied with Aspose.Cells. | Provide a step‑by‑step guide to verify that borders are preserved after copying a style in C#.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsStyleCopyDemo
{
    // This C# example creates a workbook, applies thick colored borders to cell A1, copies the complete style—including all border properties—to cell B2 using Style.Copy (or Range.CopyStyle), verifies the border colors, and saves the file as StyleCopyWithBorders.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Define a source cell with custom borders
            // -------------------------------------------------
            Cell srcCell = cells["A1"];
            srcCell.PutValue("Source");

            // Obtain the style of the source cell
            Style srcStyle = srcCell.GetStyle();

            // Set border properties on the source style
            srcStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            srcStyle.Borders[BorderType.TopBorder].Color = Color.Red;

            srcStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            srcStyle.Borders[BorderType.BottomBorder].Color = Color.Green;

            srcStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            srcStyle.Borders[BorderType.LeftBorder].Color = Color.Blue;

            srcStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            srcStyle.Borders[BorderType.RightBorder].Color = Color.Orange;

            // Apply the styled borders to the source cell
            srcCell.SetStyle(srcStyle);

            // -------------------------------------------------
            // 2. Copy the style (including borders) to a destination cell
            // -------------------------------------------------
            Cell destCell = cells["B2"];
            destCell.PutValue("Destination");

            // Option A: Use Style.Copy to duplicate the entire style object
            Style destStyle = workbook.CreateStyle();
            destStyle.Copy(srcStyle);               // Copies all attributes, borders included
            destCell.SetStyle(destStyle);           // Apply the copied style

            // Option B (alternative): Use Range.CopyStyle which also copies borders
            // Uncomment the following lines to use the range method instead of Style.Copy
            // Aspose.Cells.Range srcRange = cells.CreateRange("A1");
            // Aspose.Cells.Range destRange = cells.CreateRange("B2");
            // destRange.CopyStyle(srcRange);

            // -------------------------------------------------
            // 3. Verify that borders have been preserved
            // -------------------------------------------------
            Style verifyStyle = destCell.GetStyle();
            Console.WriteLine("Top Border Color: " + verifyStyle.Borders[BorderType.TopBorder].Color);
            Console.WriteLine("Bottom Border Color: " + verifyStyle.Borders[BorderType.BottomBorder].Color);
            Console.WriteLine("Left Border Color: " + verifyStyle.Borders[BorderType.LeftBorder].Color);
            Console.WriteLine("Right Border Color: " + verifyStyle.Borders[BorderType.RightBorder].Color);

            // -------------------------------------------------
            // 4. Save the workbook
            // -------------------------------------------------
            workbook.Save("StyleCopyWithBorders.xlsx");
        }
    }
}
