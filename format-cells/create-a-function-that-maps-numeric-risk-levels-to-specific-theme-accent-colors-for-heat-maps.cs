// Title: Map numeric risk levels to workbook theme accent colors and apply them as cell backgrounds using Aspose.Cells for .NET
// AI Prompts: Generate a C# method that accepts an Aspose.Cells Workbook and a risk value (0‑100) and returns the matching ThemeColorType accent from the workbook’s theme. | Write C# code to place a risk value into a given cell, set the cell’s solid background to the accent color returned by the risk‑to‑color function, and save the workbook to a specified file path. | Modify the tiered risk‑to‑color mapping to use custom RGB colors instead of the workbook’s built‑in theme accents while keeping the same risk ranges.
// Common Searches: asp.net map numeric risk score to workbook theme accent color using Aspose.Cells | c# set excel cell background with workbook theme colors Aspose.Cells | create heat map excel file based on risk levels Aspose.Cells for .NET | conditional formatting with theme accent colors in Aspose.Cells C# example
// Tags: risk level to theme accent color Aspose.Cells | apply workbook theme accent as cell background C# | generate heat map workbook Aspose.Cells | retrieve theme color by ThemeColorType C# | tiered risk color mapping Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Provides a helper that maps a risk score (0‑100) to one of the workbook’s theme accent colors, writes the score into a specified cell, applies the accent as a solid background, and saves the Excel file using Aspose.Cells for .NET.
public static class HeatMapHelper
{
    // Maps a numeric risk level (0‑100) to a theme accent color.
    // The workbook's theme is used to retrieve the accent colors.
    public static Color GetRiskLevelColor(Workbook workbook, double riskLevel)
    {
        // Ensure the risk level stays within the expected range.
        if (riskLevel < 0) riskLevel = 0;
        if (riskLevel > 100) riskLevel = 100;

        // Simple tiered mapping to theme accent colors.
        if (riskLevel <= 20)
            return workbook.GetThemeColor(ThemeColorType.Accent1); // low risk
        else if (riskLevel <= 40)
            return workbook.GetThemeColor(ThemeColorType.Accent2);
        else if (riskLevel <= 60)
            return workbook.GetThemeColor(ThemeColorType.Accent3);
        else if (riskLevel <= 80)
            return workbook.GetThemeColor(ThemeColorType.Accent4);
        else
            return workbook.GetThemeColor(ThemeColorType.Accent5); // high risk
    }

    // Creates a workbook, writes a risk value,
    // and applies the corresponding theme accent color as the cell background.
    public static void CreateHeatMap(string outputPath, int row, int column, double riskLevel)
    {
        // Validate output directory.
        string directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Create a new workbook (lifecycle rule).
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Write the risk level into the target cell.
        Cell cell = sheet.Cells[row, column];
        cell.PutValue(riskLevel);

        // Retrieve the appropriate accent color for the risk level.
        Color riskColor = GetRiskLevelColor(workbook, riskLevel);

        // Apply the color to the cell's style.
        Style style = cell.GetStyle();
        style.ForegroundColor = riskColor;
        style.Pattern = BackgroundType.Solid;
        cell.SetStyle(style);

        // Save the workbook (lifecycle rule).
        workbook.Save(outputPath);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Example usage: create a heat map Excel file.
            string outputPath = "HeatMapExample.xlsx";
            int row = 0;          // First row (A1)
            int column = 0;       // First column (A1)
            double riskLevel = 73.5; // Sample risk value

            HeatMapHelper.CreateHeatMap(outputPath, row, column, riskLevel);
            Console.WriteLine($"Heat map workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
