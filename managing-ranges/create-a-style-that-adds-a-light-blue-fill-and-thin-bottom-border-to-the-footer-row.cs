// Title: How to style a footer row with a light‑blue fill and thin bottom border using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply light blue background to a specific row | Add thin bottom border to a footer row with Aspose.Cells .NET | Create and assign a custom style to an entire row in a workbook using Aspose.Cells | Set solid fill color and bottom border for a row in C# Aspose.Cells | Style footer row across all columns in an Excel file with Aspose.Cells
// Tags: custom row style with background color Aspose.Cells | bottom border styling for Excel rows .NET | footer row formatting Aspose.Cells | solid fill pattern Aspose.Cells C# | apply style to all cells in a row Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The example creates a Workbook, defines a Style with a solid light‑blue fill and a thin black bottom border, and applies this style to every cell in the designated footer row (row index 9) across all columns, then saves the file as FooterStyled.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Create a style for the footer row
        Style footerStyle = workbook.CreateStyle();

        // Light blue fill
        footerStyle.ForegroundColor = Color.LightBlue;
        footerStyle.Pattern = BackgroundType.Solid;

        // Thin bottom border
        footerStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        footerStyle.Borders[BorderType.BottomBorder].Color = Color.Black;

        // Define which row is the footer (example: row 10, zero‑based index 9)
        int footerRowIndex = 9;

        // Apply the style to every cell in the footer row
        // Using the maximum column count to cover the entire row
        int maxColumn = sheet.Cells.MaxColumn;
        for (int col = 0; col < maxColumn; col++)
        {
            Cell cell = sheet.Cells[footerRowIndex, col];
            cell.SetStyle(footerStyle);
        }

        // Save the workbook (lifecycle save)
        workbook.Save("FooterStyled.xlsx");
    }
}
