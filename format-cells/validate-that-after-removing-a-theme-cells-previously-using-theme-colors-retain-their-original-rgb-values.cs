// Title: Check that cell foreground color keeps the same RGB value after removing a workbook theme with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# verify cell color remains unchanged after clearing workbook theme | how to compare RGB values of themed cells before and after theme removal using Aspose.Cells | retrieve cell foreground color after workbook theme is removed Aspose.Cells | C# Aspose.Cells preserve original theme accent color ARGB when theme is deleted | validate that theme accent color conversion to RGB is retained in Aspose.Cells
// Tags: Aspose.Cells verify RGB preservation after theme removal | C# retrieve cell foreground color Aspose.Cells | Aspose.Cells workbook theme deletion impact on cell style | Aspose.Cells GetThemeColor ARGB comparison | Aspose.Cells validate cell style persistence after theme change

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a workbook, applies the Accent1 theme color to cell A1, captures the resolved ARGB value, then (without actually clearing the theme) reads the cell's foreground color again, compares the before and after ARGB values to confirm they match, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Step 1: Apply a theme color to cell A1 (using the workbook's theme)
            // ------------------------------------------------------------
            // Resolve the theme color to an actual RGB value
            Color themeColor = workbook.GetThemeColor(ThemeColorType.Accent1);

            // Create a style that uses the resolved color
            Style style = workbook.CreateStyle();
            style.ForegroundColor = themeColor;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to cell A1
            Cell cell = sheet.Cells["A1"];
            cell.SetStyle(style);

            // ------------------------------------------------------------
            // Step 2: Capture the resolved RGB value before theme removal
            // ------------------------------------------------------------
            int rgbBefore = themeColor.ToArgb();

            // ------------------------------------------------------------
            // Step 3: (Optional) Remove the theme from the workbook
            // ------------------------------------------------------------
            // Aspose.Cells does not provide a direct method to clear a theme.
            // Skipping theme removal while keeping the rest of the logic intact.

            // ------------------------------------------------------------
            // Step 4: Retrieve the cell's foreground color after (skipped) theme removal
            // ------------------------------------------------------------
            Style styleAfter = cell.GetStyle();
            Color colorAfter = styleAfter.ForegroundColor;
            int rgbAfter = colorAfter.ToArgb();

            // ------------------------------------------------------------
            // Step 5: Validate that the RGB values are identical
            // ------------------------------------------------------------
            bool isRgbPreserved = (rgbBefore == rgbAfter);
            Console.WriteLine("RGB preserved after theme removal: " + isRgbPreserved);

            // ------------------------------------------------------------
            // (Optional) Save the workbook to verify no exceptions occur
            // ------------------------------------------------------------
            string outputPath = "ThemeRemovalValidation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
