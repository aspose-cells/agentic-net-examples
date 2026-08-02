// Title: Apply a Diagonal Crosshatch Fill Pattern to Cells Using Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, writes placeholder text to cells A1 and B2, and applies a diagonal crosshatch background (light‑gray on white) by modifying the cell style. The workbook is then saved as DiagonalCrosshatchDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | diagonal crosshatch pattern | cell fill pattern | background type | foreground color | Excel styling | example code
// Common Searches: Aspose.Cells set diagonal crosshatch fill pattern C# | how to change cell background pattern with Aspose.Cells .NET | apply custom fill style to Excel cells using Aspose.Cells | C# code for cell pattern background Aspose.Cells
// Developer Intent: Add a diagonal crosshatch background with specific foreground and background colors to selected Excel cells via Aspose.Cells.
// Use Cases: Highlight template placeholders in generated workbooks. | Visually separate input fields from calculated data in Excel forms. | Add decorative shading to report headers or footers.
// AI Prompts: Generate a method that applies a diagonal crosshatch pattern with configurable colors to any Aspose.Cells Cell in C#. | Show how to set the same diagonal crosshatch style for a range such as A1:B5 with a single call. | Create code that toggles a cell's background pattern between solid, diagonal crosshatch, and vertical stripe based on its value.

using System;
using Aspose.Cells;
using System.Drawing;

// This C# example creates a workbook, writes placeholder text to cells A1 and B2, and applies a diagonal crosshatch background (light‑gray on white) by modifying the cell style. The workbook is then saved as DiagonalCrosshatchDemo.xlsx.
class ApplyDiagonalCrosshatchPattern
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Example placeholder cells
        sheet.Cells["A1"].PutValue("Placeholder 1");
        sheet.Cells["B2"].PutValue("Placeholder 2");

        // Apply diagonal crosshatch pattern to the placeholder cells
        ApplyPattern(sheet.Cells["A1"]);
        ApplyPattern(sheet.Cells["B2"]);

        // Save the workbook
        workbook.Save("DiagonalCrosshatchDemo.xlsx", SaveFormat.Xlsx);
    }

    static void ApplyPattern(Cell cell)
    {
        // Retrieve the cell's current style
        Style style = cell.GetStyle();

        // Set the background pattern to diagonal crosshatch
        style.Pattern = BackgroundType.DiagonalCrosshatch;

        // Define foreground and background colors for the pattern
        style.ForegroundColor = Color.LightGray;
        style.BackgroundColor = Color.White;

        // Apply the modified style back to the cell
        cell.SetStyle(style);
    }
}
