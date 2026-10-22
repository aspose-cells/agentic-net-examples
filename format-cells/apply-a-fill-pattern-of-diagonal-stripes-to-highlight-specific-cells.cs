// Title: How to apply a diagonal stripe fill pattern to a cell range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set diagonal stripe background for a range of cells | Create custom cell shading with diagonal stripes using Aspose.Cells .NET | Apply style with BackgroundType.DiagonalStripe to specific cells in a workbook | Highlight B2:D4 with diagonal stripe fill pattern in C# Aspose.Cells | Using StyleFlag to apply only cell shading in Aspose.Cells
// Tags: diagonal stripe fill pattern Aspose.Cells | apply style to cell range C# | BackgroundType.DiagonalStripe usage | StyleFlag cell shading Aspose.Cells | custom cell highlighting .NET

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, defines a style with a diagonal stripe fill pattern (light blue foreground on white background), selects the B2:D4 range, applies the style to that range using a StyleFlag that enables only cell shading, and saves the file as DiagonalStripes.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with diagonal stripe fill pattern
            Style diagonalStyle = workbook.CreateStyle();
            diagonalStyle.Pattern = BackgroundType.DiagonalStripe;   // set fill pattern
            diagonalStyle.ForegroundColor = Color.LightBlue;        // stripe color
            diagonalStyle.BackgroundColor = Color.White;            // base color

            // Define which cells to highlight (e.g., B2:D4)
            AsposeRange targetRange = sheet.Cells.CreateRange("B2", "D4");

            // Apply the style to the range (only cell shading)
            StyleFlag flag = new StyleFlag();
            flag.CellShading = true;
            targetRange.ApplyStyle(diagonalStyle, flag);

            // Save the workbook
            workbook.Save("DiagonalStripes.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
