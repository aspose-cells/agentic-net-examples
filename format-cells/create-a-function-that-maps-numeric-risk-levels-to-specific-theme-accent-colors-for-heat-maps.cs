// Title: C# – Map Numeric Risk Levels to Theme Accent Colors for a Heat‑Map with Aspose.Cells
// Description: Demonstrates a static helper that converts an integer risk score into a ThemeColor (Accent1‑4) and applies the color as a solid fill style to cells in a new workbook, producing a simple risk‑heat‑map saved as RiskHeatMap.xlsx.
// Keywords: Aspose.Cells C# heat map | ThemeColor mapping | risk level color coding | ThemeColorType Accent1 | dynamic cell fill Aspose | Excel heat map example | C# numeric to color conversion
// Common Searches: Aspose.Cells map integer to theme color C# | create heat map workbook with custom colors | apply theme accent colors based on value Aspose | C# risk heat map using ThemeColor | how to color cells by numeric range Aspose.Cells
// Developer Intent: Convert numeric risk scores into specific theme accent colors and apply them to worksheet cells to generate a heat‑map workbook.
// Use Cases: Automated risk‑assessment sheets where low, medium, high, and critical values are colored green, yellow, red, and purple. | Custom visual scoring dashboards without relying on built‑in conditional formatting. | Generating a legend worksheet that pairs each risk band with its theme accent for stakeholder review.
// AI Prompts: Write a C# method that returns a ThemeColor for any numeric score, allowing a custom tint parameter. | Extend RiskHeatMapHelper to support additional risk bands and map them to Accent5‑6. | Provide code to add a legend tab that lists each risk range with its corresponding theme accent color.

using System;
using System.Drawing;
using Aspose.Cells;

namespace HeatMapExample
{
    // Helper class that maps numeric risk levels to theme accent colors
    // Demonstrates a static helper that converts an integer risk score into a ThemeColor (Accent1‑4) and applies the color as a solid fill style to cells in a new workbook, producing a simple risk‑heat‑map saved as RiskHeatMap.xlsx.
    public static class RiskHeatMapHelper
    {
        // Returns a ThemeColor based on the supplied risk level
        public static ThemeColor GetThemeColorForRisk(int riskLevel)
        {
            // Define mapping: adjust as needed for your heat‑map scale
            // Low risk (0‑2)   -> Accent1 (typically green)
            // Medium risk (3‑5) -> Accent2 (typically yellow)
            // High risk (6‑8)   -> Accent3 (typically red)
            // Very high (9+)    -> Accent4 (typically purple)
            ThemeColorType colorType;

            if (riskLevel <= 2)
                colorType = ThemeColorType.Accent1;
            else if (riskLevel <= 5)
                colorType = ThemeColorType.Accent2;
            else if (riskLevel <= 8)
                colorType = ThemeColorType.Accent3;
            else
                colorType = ThemeColorType.Accent4;

            // Tint of 0 means the original theme color; you can adjust to lighten/darken
            return new ThemeColor(colorType, 0.0);
        }
    }

    public class HeatMapGenerator
    {
        public static void CreateHeatMap()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample risk data (rows 1‑5, column A)
            int[] riskValues = { 1, 4, 7, 9, 3 };

            // Populate cells with risk values and apply theme colors
            for (int i = 0; i < riskValues.Length; i++)
            {
                int row = i + 0; // zero‑based index
                cells[row, 0].PutValue(riskValues[i]); // Column A

                // Get a ThemeColor for the current risk level
                ThemeColor themeColor = RiskHeatMapHelper.GetThemeColorForRisk(riskValues[i]);

                // Create a style that uses the theme color as the foreground (cell fill)
                Style style = workbook.CreateStyle();
                style.ForegroundThemeColor = themeColor;
                style.Pattern = BackgroundType.Solid; // Apply solid fill

                // Assign the style to the cell
                cells[row, 0].SetStyle(style);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("RiskHeatMap.xlsx");
        }

        // Entry point for demonstration
        public static void Main()
        {
            CreateHeatMap();
            Console.WriteLine("Heat map workbook created: RiskHeatMap.xlsx");
        }
    }
}
