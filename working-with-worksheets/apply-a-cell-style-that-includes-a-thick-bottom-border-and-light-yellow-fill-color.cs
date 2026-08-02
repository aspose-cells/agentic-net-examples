// Title: C# – Apply a Thick Bottom Border and Light Yellow Fill to a Cell with Aspose.Cells
// Description: Creates a new workbook, defines a style with a thick black bottom border and a solid light‑yellow background, uses a StyleFlag to apply only the border and shading, assigns the style to cell A1, and saves the file as StyledCell.xlsx.
// Keywords: Aspose.Cells | C# | .NET | cell style | bottom border | thick border | light yellow fill | background color | StyleFlag | BorderType.BottomBorder | CellBorderType.Thick | ForegroundColor | Solid fill | Excel workbook | worksheet styling
// Common Searches: Aspose.Cells add thick bottom border C# | set cell background color Aspose.Cells .NET | StyleFlag apply border and fill only | how to style a single cell with Aspose.Cells | C# example for cell border and shading in Excel
// Developer Intent: Define and apply a style that adds a thick bottom border and a light yellow fill to a specific cell using Aspose.Cells for .NET.
// Use Cases: Highlight header rows in reports with a bold bottom line and yellow shading. | Separate subtotal rows in financial sheets by applying a distinct border and fill. | Create visually consistent invoice templates where item cells have subtle fill and total rows have a prominent border.
// AI Prompts: Generate a C# snippet that applies a thick bottom border and a custom fill color to a range of cells using Aspose.Cells. | Show how to use StyleFlag to modify only the bottom border and cell shading without altering other style attributes. | Explain how to reuse a single Aspose.Cells style with different StyleFlag settings for multiple cells in a worksheet.

using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsStyleExample
{
    // Creates a new workbook, defines a style with a thick black bottom border and a solid light‑yellow background, uses a StyleFlag to apply only the border and shading, assigns the style to cell A1, and saves the file as StyledCell.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Styled Cell");

            // Create a style
            Style style = workbook.CreateStyle();

            // Set a thick bottom border (black color)
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;

            // Set light yellow fill color
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightYellow;

            // Create a style flag to apply bottom border and cell shading
            StyleFlag flag = new StyleFlag
            {
                BottomBorder = true,
                CellShading = true
            };

            // Apply the style to the cell using the flag
            cell.SetStyle(style, flag);

            // Save the workbook
            workbook.Save("StyledCell.xlsx");
        }
    }
}
