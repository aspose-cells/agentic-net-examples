using System;
using Aspose.Cells;

public class BusinessStatusThemeMapper
{
    // Maps a business status string to a specific theme accent color.
    private static ThemeColor GetThemeColorForStatus(string status)
    {
        // Define mapping between status values and theme accent types.
        // Adjust the mapping as needed for your business logic.
        switch (status?.Trim().ToUpperInvariant())
        {
            case "OPEN":
                return new ThemeColor(ThemeColorType.Accent1, 0.0); // Accent1
            case "CLOSED":
                return new ThemeColor(ThemeColorType.Accent2, 0.0); // Accent2
            case "PENDING":
                return new ThemeColor(ThemeColorType.Accent3, 0.0); // Accent3
            case "INPROGRESS":
                return new ThemeColor(ThemeColorType.Accent4, 0.0); // Accent4
            case "ONHOLD":
                return new ThemeColor(ThemeColorType.Accent5, 0.0); // Accent5
            default:
                // Fallback to a neutral accent.
                return new ThemeColor(ThemeColorType.Accent6, 0.0); // Accent6
        }
    }

    // Applies the mapped theme color as the cell's background fill.
    public static void ApplyStatusTheme(Cell cell, string status)
    {
        if (cell == null) throw new ArgumentNullException(nameof(cell));

        // Retrieve the appropriate ThemeColor based on status.
        ThemeColor themeColor = GetThemeColorForStatus(status);

        // Create a new style (using the workbook's CreateStyle method).
        Style style = cell.Worksheet.Workbook.CreateStyle();

        // Set solid fill pattern.
        style.Pattern = BackgroundType.Solid;

        // Apply the theme color to the cell's background.
        style.BackgroundThemeColor = themeColor;

        // Assign the style to the cell.
        cell.SetStyle(style);
    }

    // Demonstrates usage: creates a workbook, fills status values, and applies theme colors.
    public static void RunDemo()
    {
        // Create a new workbook (lifecycle rule: create).
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data.
        string[] statuses = { "Open", "Closed", "Pending", "InProgress", "OnHold", "Unknown" };

        // Populate column A with status values and apply theme colors.
        for (int i = 0; i < statuses.Length; i++)
        {
            Cell cell = sheet.Cells[i, 0]; // Column A
            cell.PutValue(statuses[i]);
            ApplyStatusTheme(cell, statuses[i]);
        }

        // Save the workbook (lifecycle rule: save).
        workbook.Save("BusinessStatusTheme.xlsx");
    }
}

// Entry point for testing.
class Program
{
    static void Main()
    {
        BusinessStatusThemeMapper.RunDemo();
        Console.WriteLine("Workbook created with status-themed cell fills.");
    }
}