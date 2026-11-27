// Title: C# function to get a workbook’s theme accent color for a given risk level with Aspose.Cells
// Common Searches: Aspose.Cells C# get theme accent color based on risk level value | how to map numeric risk levels to theme colors in a heat map using Aspose.Cells | C# retrieve workbook theme accent color for heat‑map risk categories | using Aspose.Cells to assign heat map colors from theme accents in C# | example code for getting theme accent colors from a workbook in Aspose.Cells
// Tags: retrieve theme accent color Aspose.Cells C# | risk level color mapping workbook theme | heat map accent color selection Aspose.Cells | theme color lookup by numeric index Aspose.Cells | C# get workbook theme color for heat map

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The GetRiskLevelColor method accepts an Aspose.Cells Workbook and an integer risk level (1‑5) and returns the corresponding theme accent color (Accent1‑Accent5) from the workbook, or Transparent for null workbooks or out‑of‑range values.
public static class HeatMapHelper
{
    /// <param name="workbook">The Aspose.Cells workbook containing the theme.</param>
    /// <param name="riskLevel">Numeric risk level (e.g., 1‑5).</param>
    /// <returns>The corresponding theme accent color; Transparent if out of range.</returns>
    public static Color GetRiskLevelColor(Workbook workbook, int riskLevel)
    {
        // Ensure the workbook is not null to avoid NullReferenceException.
        if (workbook == null)
            return Color.Transparent;

        // Map risk levels to theme accent colors.
        // Adjust the mapping as needed for your specific heat‑map design.
        switch (riskLevel)
        {
            case 1:
                return workbook.GetThemeColor(ThemeColorType.Accent1);
            case 2:
                return workbook.GetThemeColor(ThemeColorType.Accent2);
            case 3:
                return workbook.GetThemeColor(ThemeColorType.Accent3);
            case 4:
                return workbook.GetThemeColor(ThemeColorType.Accent4);
            case 5:
                return workbook.GetThemeColor(ThemeColorType.Accent5);
            default:
                // Return Transparent for undefined risk levels.
                return Color.Transparent;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Path to an optional template workbook.
            string templatePath = "template.xlsx";

            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one.
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Example usage: retrieve the color for risk level 3.
            Color riskColor = HeatMapHelper.GetRiskLevelColor(workbook, 3);
            Console.WriteLine($"Risk level 3 color: {riskColor}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
