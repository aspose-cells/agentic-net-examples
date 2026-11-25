// Title: How to add a thick left border and light‑blue fill to a cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set thick left border on a single cell | C# apply light blue solid background to a specific cell using Aspose.Cells | Styling a cell with border and fill in Aspose.Cells .NET example | Create custom cell style with left border and background color in Aspose.Cells | Set cell style programmatically in Aspose.Cells workbook C#
// Tags: left border styling Aspose.Cells | solid light blue fill Aspose.Cells | cell style definition Aspose.Cells | border with fill Aspose.Cells | save workbook with styled cell Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The example creates a workbook, defines a style that adds a thick left border and a light‑blue solid background, applies the style to cell A1, and saves the file as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the target cell (e.g., A1)
        Cell cell = sheet.Cells["A1"];

        // Create a new style
        Style style = workbook.CreateStyle();

        // Apply a thick left border
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
        style.Borders[BorderType.LeftBorder].Color = Color.Black;

        // Set a light blue background fill
        style.ForegroundColor = Color.LightBlue;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to the cell
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("StyledCell.xlsx");
    }
}
