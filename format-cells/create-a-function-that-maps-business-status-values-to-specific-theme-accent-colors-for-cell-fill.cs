// Title: How to map business status strings to theme accent fill colors for Excel cells using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells set cell background color based on status value | apply theme accent colors to Excel cells with Aspose.Cells .NET | map status text to solid fill color in a worksheet using Aspose.Cells | conditional cell styling by status string in Aspose.Cells for .NET | save Excel file with status‑based cell colors using Aspose.Cells C#
// Tags: status string to theme accent color Aspose.Cells | apply solid fill background to worksheet cell C# | custom color mapping for Excel cell styling .NET | set cell style foreground and background Aspose.Cells | write status text and color cell Aspose.Cells C# | save workbook with colored status cells .NET

using System;
using System.Drawing;
using Aspose.Cells;

// This example defines a GetColorForStatus method that converts business status strings (Open, Closed, Pending, In Progress, Canceled) into specific theme accent ARGB colors, and an ApplyStatusFill method that writes the status into a cell, applies a solid background fill with the mapped color and white font using Aspose.Cells style objects. The Main routine demonstrates creating a workbook, applying the status‑based fills to several cells, and saving the result as an Excel file.
public class StatusColorMapper
{
    // Maps a business status string to a specific solid fill color.
    private static Color? GetColorForStatus(string status)
    {
        if (string.IsNullOrWhiteSpace(status))
            return null;

        switch (status.Trim().ToLower())
        {
            case "open":
                return Color.FromArgb(0, 112, 192);   // Blue (Accent 1)
            case "closed":
                return Color.FromArgb(255, 0, 0);     // Red (Accent 2)
            case "pending":
                return Color.FromArgb(255, 192, 0);   // Orange (Accent 3)
            case "in progress":
                return Color.FromArgb(0, 176, 80);    // Green (Accent 4)
            case "canceled":
                return Color.FromArgb(112, 48, 160);  // Purple (Accent 5)
            default:
                return null;                         // No specific mapping
        }
    }

    // Applies the mapped solid fill color to the specified cell.
    public static void ApplyStatusFill(Worksheet sheet, string cellName, string status)
    {
        // Write the status text into the cell.
        var cell = sheet.Cells[cellName];
        cell.PutValue(status);

        // Retrieve the color for the given status.
        var fillColor = GetColorForStatus(status);
        if (!fillColor.HasValue)
            return; // No color mapping; leave default formatting.

        // Configure the cell style with a solid background using the mapped color.
        var style = cell.GetStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = fillColor.Value;

        // Optional: set font color for better contrast.
        style.Font.Color = Color.White;

        // Apply the style back to the cell.
        cell.SetStyle(style);
    }

    // Example entry point demonstrating usage.
    public static void Main()
    {
        try
        {
            // Create a new workbook.
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Apply status-based fill colors to a range of cells.
            ApplyStatusFill(sheet, "A1", "Open");
            ApplyStatusFill(sheet, "A2", "Closed");
            ApplyStatusFill(sheet, "A3", "Pending");
            ApplyStatusFill(sheet, "A4", "In Progress");
            ApplyStatusFill(sheet, "A5", "Canceled");
            ApplyStatusFill(sheet, "A6", "Other"); // No mapping; default style.

            // Save the workbook.
            workbook.Save("StatusColors.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
