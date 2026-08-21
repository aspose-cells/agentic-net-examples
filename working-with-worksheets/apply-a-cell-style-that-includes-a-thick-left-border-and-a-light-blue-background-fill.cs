// Title: Style a Cell with Thick Left Border and Light‑Blue Fill using Aspose.Cells for .NET
// Description: Creates a workbook, selects cell B2, and applies a Style with a solid light‑blue background and a thick left border. A StyleFlag limits the changes to the left border and cell shading before saving as StyledCell.xlsx.
// Keywords: Aspose.Cells cell style | C# thick left border | light blue fill Aspose.Cells | StyleFlag selective styling | solid background color .NET | border formatting Aspose.Cells
// Common Searches: Aspose.Cells add thick left border to a cell | set light blue background for a specific cell Aspose.Cells | how to use StyleFlag for border only in Aspose.Cells | apply custom cell style without affecting other borders .NET
// Developer Intent: Apply a thick left border and a light‑blue background to a single cell using Aspose.Cells for .NET.
// Use Cases: Highlight header cells with a colored background and a distinct left separator. | Create a visual column divider by styling cells with a prominent left border while leaving other borders unchanged. | Emphasize rows in a financial report with light‑blue shading and a bold left edge.
// AI Prompts: Generate C# code with Aspose.Cells that styles a range of cells using a thick left border and light‑blue fill. | Explain the purpose of StyleFlag in Aspose.Cells and how it can apply only border and shading attributes. | Show how to set different border colors while keeping a solid fill using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, selects cell B2, and applies a Style with a solid light‑blue background and a thick left border. A StyleFlag limits the changes to the left border and cell shading before saving as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Choose a cell to style and put a sample value
        Cell cell = cells["B2"];
        cell.PutValue("Styled Cell");

        // Create a style object
        Style style = workbook.CreateStyle();

        // Set light blue background fill
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightBlue;

        // Configure a thick left border
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
        style.Borders[BorderType.LeftBorder].Color = Color.Black; // border color (optional)

        // Create a style flag to apply left border and cell shading
        StyleFlag flag = new StyleFlag
        {
            LeftBorder = true,
            CellShading = true
        };

        // Apply the style to the cell using the flag
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("StyledCell.xlsx");
    }
}
