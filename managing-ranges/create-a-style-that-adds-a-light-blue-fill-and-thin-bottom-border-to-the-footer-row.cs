// Title: Aspose.Cells for .NET – Apply Light Blue Fill and Thin Bottom Border to a Footer Row
// Description: Demonstrates how to create a workbook, locate the footer row, build a Style with a solid LightBlue background and a thin black bottom border, apply it to the row, and save the file as FooterStyle.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells style row background | C# light blue fill Aspose.Cells | thin bottom border Aspose.Cells | footer row formatting .NET | apply style to worksheet row | Aspose.Cells GitHub example | Excel row styling C#
// Common Searches: how to set row background color in Aspose.Cells | add thin bottom border to a row using Aspose.Cells .NET | style footer row in Excel with Aspose.Cells | Aspose.Cells C# example for row styling | apply solid fill to a specific row in Aspose.Cells
// Developer Intent: Add a light‑blue fill and a thin bottom border to the worksheet’s footer row programmatically.
// Use Cases: Highlight the total or summary row in financial reports with a distinct background and border. | Create a printable invoice where the final row stands out for quick visual reference. | Standardize footer appearance across multiple sheets in a workbook by reusing a single Style object.
// AI Prompts: Generate C# code with Aspose.Cells that styles a footer row using a LightBlue solid fill and a thin black bottom border. | Provide a reusable method that accepts a Worksheet and row index, then applies a light blue background and thin bottom border. | Explain how to clone a Style in Aspose.Cells and apply it to several footer rows within the same workbook.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsFooterStyle
{
    // Demonstrates how to create a workbook, locate the footer row, build a Style with a solid LightBlue background and a thin black bottom border, apply it to the row, and save the file as FooterStyle.xlsx using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the footer row index (for example, one row after the last used row)
            int footerRowIndex = worksheet.Cells.MaxDataRow + 1;

            // Create a new style
            Style style = workbook.CreateStyle();

            // Set light blue fill
            style.BackgroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;

            // Set a thin bottom border
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;

            // Apply the style to the footer row
            Row footerRow = worksheet.Cells.Rows[footerRowIndex];
            footerRow.SetStyle(style);

            // Save the workbook
            workbook.Save("FooterStyle.xlsx");
        }
    }
}
