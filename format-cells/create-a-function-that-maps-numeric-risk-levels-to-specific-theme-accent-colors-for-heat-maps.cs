// Title: Map Risk Levels to Excel Theme Accent Colors and Build a Heat‑Map with Aspose.Cells (C#)
// Description: Shows how to clamp a numeric risk level (1‑5) to a ThemeColor using ThemeColorType.Accent1‑Accent5, apply the color to cell styles, and generate a simple risk heat‑map workbook saved as RiskHeatMap.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | ThemeColor | Accent colors | risk heat map | Excel theme colors | cell background color | heat map generation | ThemeColorType | range clamping | Excel styling
// Common Searches: Aspose.Cells map numeric value to theme accent color | Create heat map in Excel using C# Aspose.Cells | Set cell background based on risk level Aspose.Cells | Clamp values before applying ThemeColor in Aspose.Cells | How to use ThemeColorType.Accent in Aspose.Cells
// Developer Intent: Create an Excel workbook where each cell’s background reflects a risk score using built‑in theme accent colors.
// Use Cases: Visualize financial risk matrices with gradient colors. | Generate project‑risk dashboards directly from .NET applications. | Integrate risk‑based color coding into automated reporting pipelines. | Apply custom theme‑based palettes for branding or compliance.
// AI Prompts: Write a C# method that receives an integer 1‑5 and returns a ThemeColor using ThemeColorType.Accent1‑Accent5, including bounds checking. | Provide code that reads a two‑dimensional risk array, writes the values to an Aspose.Cells worksheet, sets each cell’s ForegroundThemeColor via the risk‑to‑theme mapper, and saves the file. | Explain how to replace the default Accent colors with a user‑defined RGB palette while keeping the same risk‑level mapping. | Show how to extend the mapper to support additional risk levels and custom tint adjustments.

using System;
using Aspose.Cells;
using System.Drawing;

// Shows how to clamp a numeric risk level (1‑5) to a ThemeColor using ThemeColorType.Accent1‑Accent5, apply the color to cell styles, and generate a simple risk heat‑map workbook saved as RiskHeatMap.xlsx with Aspose.Cells for .NET.
public class HeatMapRiskColorMapper
{
    // Maps a numeric risk level (1‑5) to a ThemeColor using Accent theme colors.
    public static ThemeColor GetRiskThemeColor(int riskLevel)
    {
        // Clamp risk level to the expected range.
        if (riskLevel < 1) riskLevel = 1;
        if (riskLevel > 5) riskLevel = 5;

        ThemeColorType type = ThemeColorType.Accent1; // default
        double tint = 0.0; // no tint adjustment

        switch (riskLevel)
        {
            case 1:
                type = ThemeColorType.Accent1; // low risk (e.g., green)
                break;
            case 2:
                type = ThemeColorType.Accent2; // low‑medium
                break;
            case 3:
                type = ThemeColorType.Accent3; // medium
                break;
            case 4:
                type = ThemeColorType.Accent4; // high‑medium
                break;
            case 5:
                type = ThemeColorType.Accent5; // high risk (e.g., red)
                break;
        }

        return new ThemeColor(type, tint);
    }

    // Creates a simple heat‑map workbook where each cell background reflects its risk level.
    public static void CreateRiskHeatMap()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample risk data (rows x columns).
        int[,] riskData = new int[,]
        {
            {1, 2, 3},
            {2, 3, 4},
            {3, 4, 5}
        };

        // Populate cells and apply theme‑based foreground colors.
        for (int row = 0; row < riskData.GetLength(0); row++)
        {
            for (int col = 0; col < riskData.GetLength(1); col++)
            {
                int risk = riskData[row, col];
                string address = CellsHelper.CellIndexToName(row, col);
                cells[address].PutValue(risk);

                // Create a style that uses the mapped ThemeColor.
                Style style = workbook.CreateStyle();
                style.ForegroundThemeColor = GetRiskThemeColor(risk);
                style.Pattern = BackgroundType.Solid;

                cells[address].SetStyle(style);
            }
        }

        // Save the workbook.
        workbook.Save("RiskHeatMap.xlsx");
    }

    // Entry point.
    public static void Main()
    {
        CreateRiskHeatMap();
    }
}
