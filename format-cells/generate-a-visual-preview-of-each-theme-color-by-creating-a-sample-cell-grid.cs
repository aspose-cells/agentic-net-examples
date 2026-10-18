// Title: Generate an Excel workbook that displays a preview grid of all theme colors using Aspose.Cells for .NET (C#)
// Common Searches: how to create a theme color preview sheet with Aspose.Cells in C# | sample code for displaying Excel theme colors as cell backgrounds using Aspose.Cells | C# Aspose.Cells example to list Accent1‑Accent6 and hyperlink theme colors | auto‑fit columns after applying theme color styles in Aspose.Cells workbook | retrieve and use workbook GetThemeColor method in Aspose.Cells .NET
// Tags: aspocells theme color preview grid | c# workbook generate theme color samples | aspocells getthemecolor method usage | excel theme color visualization c# | aspocells auto-fit columns after styling

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, adds a header row, and iterates through the ThemeColorType values (Accent1‑6, Hyperlink, FollowedHyperlink). For each theme color it writes the name, obtains the actual Color via workbook.GetThemeColor, applies it as a solid background to a sample cell, auto‑fits the columns, and saves the file as ThemeColorsPreview.xlsx, with exception handling.
class ThemeColorsPreview
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the theme colors to preview (exclude unavailable Dark/Light types)
            ThemeColorType[] themeColors = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6,
                ThemeColorType.Hyperlink,
                ThemeColorType.FollowedHyperlink
            };

            // Header row
            sheet.Cells[0, 0].PutValue("Theme Color");
            sheet.Cells[0, 1].PutValue("Sample");

            // Apply bold style to header cells
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            StyleFlag headerFlag = new StyleFlag() { All = true };
            sheet.Cells[0, 0].SetStyle(headerStyle, headerFlag);
            sheet.Cells[0, 1].SetStyle(headerStyle, headerFlag);

            // Populate each theme color with a sample cell
            for (int i = 0; i < themeColors.Length; i++)
            {
                int row = i + 1;

                // Write the name of the theme color
                sheet.Cells[row, 0].PutValue(themeColors[i].ToString());

                // Create a style using the theme color as background
                Style sampleStyle = workbook.CreateStyle();
                Color themeColor = workbook.GetThemeColor(themeColors[i]);
                sampleStyle.ForegroundColor = themeColor;
                sampleStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the sample cell
                StyleFlag sampleFlag = new StyleFlag() { All = true };
                sheet.Cells[row, 1].SetStyle(sampleStyle, sampleFlag);
            }

            // Auto-fit columns for better visibility
            sheet.AutoFitColumn(0);
            sheet.AutoFitColumn(1);

            // Save the workbook
            string outputPath = "ThemeColorsPreview.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
