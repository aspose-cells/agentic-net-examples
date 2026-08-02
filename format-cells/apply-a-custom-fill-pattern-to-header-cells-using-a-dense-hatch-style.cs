// Title: Apply a Dense Hatch Fill Pattern to Header Cells with Aspose.Cells for .NET
// Description: Shows how to create a workbook, write header values, define a style that uses the ThickDiagonalCrosshatch pattern with LightGray foreground and DarkGray background, enable cell shading, apply the style to the first‑row range, and save the file as HeaderPattern.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | .NET | Excel fill pattern | ThickDiagonalCrosshatch | header cell style | cell shading | BackgroundType | foreground color | background color | custom pattern
// Common Searches: Aspose.Cells set hatch pattern for cells | C# apply ThickDiagonalCrosshatch to a range | how to use BackgroundType in Aspose.Cells | custom header style Excel Aspose | cell shading with foreground and background colors
// Developer Intent: Add a dense hatch background to a header row in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create printable reports where header rows stand out with a distinctive hatch pattern. | Separate worksheet sections visually by applying custom patterns to specific ranges. | Automate Excel generation with styled headers that improve readability and hierarchy.
// AI Prompts: Generate C# code that applies a ThickDiagonalCrosshatch pattern with LightGray foreground and DarkGray background to a specified range using Aspose.Cells. | Show how to switch the hatch style and colors of a header row style in Aspose.Cells for .NET. | Provide an example that applies the same dense hatch pattern to header rows across multiple worksheets in a single workbook. | Write a function that receives a range address and pattern colors, then applies a custom fill pattern with Aspose.Cells.

using System.Drawing;
using Aspose.Cells;

// Shows how to create a workbook, write header values, define a style that uses the ThickDiagonalCrosshatch pattern with LightGray foreground and DarkGray background, enable cell shading, apply the style to the first‑row range, and save the file as HeaderPattern.xlsx using Aspose.Cells in C#.
class ApplyHeaderFillPattern
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate header cells (first row, columns A to D)
        for (int col = 0; col < 4; col++)
        {
            Cell cell = sheet.Cells[0, col];
            cell.PutValue($"Header {col + 1}");
        }

        // Create a style for the header cells
        Style headerStyle = workbook.CreateStyle();

        // Use a dense hatch pattern (ThickDiagonalCrosshatch)
        headerStyle.Pattern = BackgroundType.ThickDiagonalCrosshatch;

        // Define foreground and background colors for the pattern
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.BackgroundColor = Color.DarkGray;

        // Apply the style to the header range
        StyleFlag flag = new StyleFlag();
        flag.CellShading = true; // Enable pattern application
        sheet.Cells.CreateRange(0, 0, 1, 4).ApplyStyle(headerStyle, flag);

        // Save the workbook
        workbook.Save("HeaderPattern.xlsx");
    }
}
