// Title: Apply a Thick Left Border and Light‑Blue Fill to a Cell with Aspose.Cells for .NET
// Description: Demonstrates how to create a style with a solid light‑blue background and a thick left border, apply it to a single cell using a StyleFlag, and save the workbook as StyledCell.xlsx.
// Keywords: Aspose.Cells style left border | C# cell background color | thick left border Aspose | light blue fill Aspose.Cells | StyleFlag partial styling | apply cell style .NET | Excel cell formatting Aspose
// Common Searches: Aspose.Cells add thick left border to a cell | set light blue background for a cell in C# | partial cell styling with StyleFlag Aspose | how to apply border and fill to one cell Aspose.Cells | C# Excel cell style left border only
// Developer Intent: The developer wants to format a single cell with a thick left border and a light‑blue background using Aspose.Cells for .NET.
// Use Cases: Highlight header or subtotal rows with a distinct left border and blue shading for better visual separation. | Create a data‑entry column where each cell is visually isolated by a thick left border while sharing a uniform background color. | Design a financial report that uses blue‑filled cells with prominent left borders to draw attention to key figures.
// AI Prompts: Generate code to apply a thick left border and light‑blue fill to an entire column in Aspose.Cells for .NET. | Show how to conditionally add a left border only when a cell's numeric value exceeds a threshold, keeping the blue background. | Explain how to reuse a Style and StyleFlag across multiple worksheets to improve performance when applying identical border and fill settings.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a style with a solid light‑blue background and a thick left border, apply it to a single cell using a StyleFlag, and save the workbook as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Create a new style
        Style style = workbook.CreateStyle();

        // Set a solid fill with light blue background
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightBlue;

        // Configure the left border: thick line
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
        style.Borders[BorderType.LeftBorder].Color = Color.Black; // optional border color

        // Create a style flag to apply only the left border and cell shading
        StyleFlag flag = new StyleFlag
        {
            LeftBorder = true,
            CellShading = true
        };

        // Apply the style to a specific cell (e.g., B2)
        Cell targetCell = cells["B2"];
        targetCell.PutValue("Styled Cell");
        targetCell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("StyledCell.xlsx");
    }
}
