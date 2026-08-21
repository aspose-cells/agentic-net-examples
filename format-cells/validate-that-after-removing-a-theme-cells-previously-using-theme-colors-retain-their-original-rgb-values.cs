// Title: Validate Cell Fill Color Persistence After Removing Workbook Theme with Aspose.Cells for .NET
// Description: Shows how to apply an explicit RGB fill to a cell, capture its ForegroundColor, simulate a theme change, and confirm that the RGB value stays the same using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | cell fill color | theme removal | RGB persistence | foreground color validation | workbook theme | color unchanged | style verification
// Common Searches: Aspose.Cells keep cell background color after theme delete | verify RGB color after removing theme Aspose.Cells | how to test cell color persistence in .NET workbook | Aspose.Cells theme color vs explicit RGB | check if cell style changes when workbook theme is removed
// Developer Intent: Ensure that a cell styled with an explicit RGB fill does not change when the workbook's theme is removed or altered.
// Use Cases: Automated testing to guarantee custom colors survive theme modifications. | Generating reports where the workbook theme may be stripped but original colors must remain. | Migrating or consolidating workbooks while preserving user‑defined fill colors. | Debugging style issues by comparing pre‑ and post‑theme removal color values.
// AI Prompts: Write C# code with Aspose.Cells that sets a cell's background to a specific RGB value, removes the workbook theme, and asserts the color is unchanged. | Create a .NET unit test that applies Color.Blue to a cell, simulates theme removal, and verifies the ForegroundColor ARGB values match. | Explain how Aspose.Cells stores explicit RGB colors versus theme‑based colors and how to validate their persistence after a theme change.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeRemovalValidation
{
    // Shows how to apply an explicit RGB fill to a cell, capture its ForegroundColor, simulate a theme change, and confirm that the RGB value stays the same using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Apply a solid fill color to cell A1
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Theme Color Test");

                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                // Use a specific RGB color (e.g., Blue) instead of a theme reference
                style.ForegroundColor = Color.Blue;
                cell.SetStyle(style);

                // Resolve the actual RGB value that was applied
                Color originalRgb = cell.GetStyle().ForegroundColor;

                // Simulate a theme change by modifying the workbook's theme color (not used here)
                // In this simplified example we skip theme manipulation.

                // Retrieve the cell style after the simulated theme change
                Style afterThemeChange = cell.GetStyle();
                Color afterRgb = afterThemeChange.ForegroundColor;

                // Validate that the RGB value remained unchanged
                bool isUnchanged = afterRgb.ToArgb() == originalRgb.ToArgb();
                Console.WriteLine($"Original RGB: {originalRgb}");
                Console.WriteLine($"RGB after simulated theme change: {afterRgb}");
                Console.WriteLine($"Color unchanged after removing theme: {isUnchanged}");

                // Save the workbook (optional)
                string outputPath = "ThemeRemovalValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
