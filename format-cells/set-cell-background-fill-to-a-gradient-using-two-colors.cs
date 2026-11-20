// Title: Set a two‑color gradient‑like background for an Excel cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to apply start and end colors to a cell background | simulate gradient fill in Excel cell with Aspose.Cells .NET | Aspose.Cells set foreground and background colors for a single cell | apply solid fill pattern with custom colors to Excel cell using Aspose.Cells | create gradient‑style cell shading in Aspose.Cells C# example
// Tags: Aspose.Cells set cell foreground and background colors | solid fill pattern gradient simulation Aspose.Cells | StyleFlag cell shading C# Aspose.Cells | Excel cell background custom colors Aspose.Cells .NET | save workbook with styled cell Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, writes text to cell A1, defines a style with Yellow as the foreground (start) color and Orange as the background (end) color, sets the fill pattern to solid (to simulate a gradient), applies the style to the cell using a StyleFlag that enables cell shading, and saves the workbook as GradientFill.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Gradient Fill (simulated with solid colors)");

            // Create a new style
            Style style = workbook.CreateStyle();

            // Set start (foreground) and end (background) colors
            style.ForegroundColor = Color.Yellow;   // start color
            style.BackgroundColor = Color.Orange;   // end color

            // Apply solid fill pattern (gradient not directly supported in this version)
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the cell (only shading part)
            StyleFlag flag = new StyleFlag { CellShading = true };
            cell.SetStyle(style, flag);

            // Save the workbook
            string outputPath = "GradientFill.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
