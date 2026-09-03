// Title: How to add a thick bottom border and light‑yellow solid fill to a cell using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a Style in Aspose.Cells, sets the BottomBorder line style to Thick, assigns a light‑yellow solid fill, applies the style to cell A1, and saves the workbook. | Show the steps to configure a cell’s border and background color in Aspose.Cells for .NET, including setting BorderType.BottomBorder, CellBorderType.Thick, and a light‑yellow ForegroundColor.
// Common Searches: Aspose.Cells C# set thick bottom border for a specific cell | How to fill a cell with light yellow color using Aspose.Cells .NET | C# example applying border and background style to a cell in Aspose.Cells | Aspose.Cells style object bottom border line style Thick | Create a styled cell with solid fill and border in Aspose.Cells for .NET
// Tags: Aspose.Cells configure bottom border style | Aspose.Cells set cell fill color yellow | C# Aspose.Cells create cell style | Aspose.Cells border and background formatting | Aspose.Cells workbook cell styling example

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates creating a workbook in C#, defining a Style with a thick bottom border and a light‑yellow solid fill, applying the style to cell A1, and saving the file as StyledCell.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the target cell (e.g., A1)
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample");

        // Create a new style object
        Style style = workbook.CreateStyle();

        // Apply a thick bottom border
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
        style.Borders[BorderType.BottomBorder].Color = Color.Black; // optional color

        // Set a light yellow fill color
        style.ForegroundColor = Color.FromArgb(255, 255, 204); // light yellow
        style.Pattern = BackgroundType.Solid;

        // Assign the style to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("StyledCell.xlsx");
    }
}
