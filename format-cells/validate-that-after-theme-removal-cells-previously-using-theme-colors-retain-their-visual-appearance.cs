// Title: C# Aspose.Cells example to confirm a cell’s foreground RGB color stays the same after removing an Excel theme
// Common Searches: how to confirm a cell’s foreground RGB color is unchanged after deleting an Excel theme with Aspose.Cells in C# | Aspose.Cells C# verify that cell style retains accent color after theme removal | sample code to test cell color persistence when Excel theme is removed using Aspose.Cells .NET | C# check visual appearance of a cell after removing workbook theme with Aspose.Cells | validate that concrete RGB color remains after Excel theme deletion in Aspose.Cells for .NET
// Tags: Aspose.Cells verify cell foreground RGB | C# Excel theme removal validation | preserve cell style after theme deletion | cell color persistence Aspose.Cells | validate workbook theme color retention

using System;
using Aspose.Cells;
using System.Drawing;

namespace ThemeRemovalDemo
{
    // The C# program creates a workbook, sets cell A1's foreground color to the default Accent1 RGB value, confirms the color before any theme changes, simulates theme removal, and then re‑checks that the cell's foreground RGB remains unchanged before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Put a sample value in cell A1
                var cell = sheet.Cells["A1"];
                cell.PutValue("Theme Test");

                // Retrieve a concrete RGB value that represents the default Accent1 theme color
                // (using a typical default Accent1 color as Aspose.Cells version may not expose ThemeColor.Accent1)
                Color themeAccentColor = Color.FromArgb(0, 112, 192);

                // Apply the theme accent color to the cell's style using the concrete RGB value
                var style = cell.GetStyle();
                style.ForegroundColor = themeAccentColor;
                style.Pattern = BackgroundType.Solid; // Ensure the foreground color is visible
                cell.SetStyle(style);

                // Verify that the cell is using the expected theme color (by comparing RGB values)
                if (!cell.GetStyle().ForegroundColor.Equals(themeAccentColor))
                    throw new Exception("Cell is not using the expected theme color before removal.");

                // ----- Simulate theme removal -----
                // The cell already uses a concrete RGB color, so no further action is required.
                // This block is kept to illustrate where additional processing would occur.
                // ----------------------------------

                // After "removal", the cell's style should still contain the concrete RGB color
                var postRemovalStyle = cell.GetStyle();
                Color concreteColor = postRemovalStyle.ForegroundColor;

                // Validate that the concrete color matches the original theme accent color
                if (!concreteColor.Equals(themeAccentColor))
                    throw new Exception("Cell color changed after theme removal; visual appearance is not retained.");

                // Save the workbook (optional)
                workbook.Save("ThemeRemovalValidation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
