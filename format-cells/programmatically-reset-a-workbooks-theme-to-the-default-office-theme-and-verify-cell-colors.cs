// Title: Reset an Excel workbook to the default Office theme and verify cell background color using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# reset workbook theme to default Office theme | verify cell background color after resetting Excel theme with Aspose.Cells | remove custom cell style and apply default style in Aspose.Cells .NET | programmatically clear workbook theme using Aspose.Cells and check default formatting
// Tags: reset workbook theme Aspose.Cells | apply default style to cell Aspose.Cells | retrieve cell background color Aspose.Cells | default Office theme Excel Aspose.Cells | clear custom cell formatting Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a new workbook, applies a LightBlue solid fill to cell A1, saves the file, then simulates a theme reset by assigning the workbook's DefaultStyle to the cell. It reads the resulting cell style to obtain the foreground color, determines whether the background matches the default (transparent or white), and saves the workbook after the simulated reset.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data and apply a custom background color to a cell
            Worksheet ws = wb.Worksheets[0];
            Cell cell = ws.Cells["A1"];
            cell.PutValue("Sample");

            Style customStyle = wb.CreateStyle();
            customStyle.ForegroundColor = Color.LightBlue;
            customStyle.Pattern = BackgroundType.Solid;
            cell.SetStyle(customStyle);

            // Save the workbook before attempting any theme changes (optional)
            string beforePath = "BeforeReset.xlsx";
            wb.Save(beforePath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(beforePath)}");

            // Simulate a theme reset by clearing the cell style (reverting to default)
            cell.SetStyle(wb.DefaultStyle);

            // Verify the cell's background color after the simulated theme reset
            Style resultStyle = cell.GetStyle();
            Color bgColor = resultStyle.ForegroundColor;

            Console.WriteLine("Cell A1 background color after theme reset simulation: " + bgColor.ToArgb());

            // The default background is typically Transparent (no fill) or White.
            bool isDefault = bgColor.IsEmpty || bgColor.ToArgb() == Color.White.ToArgb();
            Console.WriteLine("Is default background: " + isDefault);

            // Save the workbook after the theme simulation
            string afterPath = "AfterReset.xlsx";
            wb.Save(afterPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(afterPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
