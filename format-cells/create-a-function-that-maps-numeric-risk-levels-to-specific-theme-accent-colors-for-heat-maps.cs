// Title: Map numeric risk levels to Excel theme accent colors and apply cell background colors with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set cell background color based on numeric risk level | How to create a risk heat map in Excel using Aspose.Cells .NET | Apply conditional color formatting to cells programmatically with Aspose.Cells | Map risk score ranges to specific colors in an Aspose.Cells workbook | Set solid background pattern for Excel cells using custom color logic in C#
// Tags: risk level to color mapping Aspose.Cells | apply solid background color to Excel cell C# | conditional cell styling based on numeric value Aspose.Cells | theme accent color heat map Aspose.Cells | programmatic cell formatting Excel .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example defines a GetRiskColor method that returns a System.Drawing.Color for a given risk score range and an ApplyRiskColor method that writes the score to a cell, sets the cell's foreground color using the mapping, and applies a solid background pattern. The Main routine creates a workbook, colors cells A1‑A5 with different risk levels, and saves the file as RiskHeatMap.xlsx.
public static class RiskColorMapper
{
    /// <param name="riskLevel">Numeric risk level (e.g., 0‑100).</param>
    /// <returns>Corresponding System.Drawing.Color.</returns>
    public static Color GetRiskColor(double riskLevel)
    {
        // Low risk (0‑20) → Green
        if (riskLevel <= 20)
            return Color.Green;

        // Moderate risk (21‑40) → Teal
        if (riskLevel <= 40)
            return Color.Teal;

        // Elevated risk (41‑60) → Orange
        if (riskLevel <= 60)
            return Color.Orange;

        // High risk (61‑80) → Red
        if (riskLevel <= 80)
            return Color.Red;

        // Critical risk (81‑100) → Purple
        return Color.Purple;
    }

    // Example usage: apply the mapped color to a cell for a heat‑map effect.
    public static void ApplyRiskColor(Worksheet sheet, string cellName, double riskLevel)
    {
        Cell cell = sheet.Cells[cellName];
        cell.PutValue(riskLevel);

        Style style = cell.GetStyle();
        style.ForegroundColor = GetRiskColor(riskLevel);
        style.Pattern = BackgroundType.Solid; // Ensure background is filled
        cell.SetStyle(style);
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Apply risk colors to a few sample cells.
            RiskColorMapper.ApplyRiskColor(sheet, "A1", 10); // Low risk
            RiskColorMapper.ApplyRiskColor(sheet, "A2", 30); // Moderate risk
            RiskColorMapper.ApplyRiskColor(sheet, "A3", 50); // Elevated risk
            RiskColorMapper.ApplyRiskColor(sheet, "A4", 70); // High risk
            RiskColorMapper.ApplyRiskColor(sheet, "A5", 90); // Critical risk

            // Define output path.
            string outputPath = "RiskHeatMap.xlsx";

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
