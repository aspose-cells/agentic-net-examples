using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample business status values
        string[] statuses = { "Open", "Closed", "Pending", "InProgress", "Cancelled", "Unknown" };

        // Populate cells with status values and apply theme accent colors
        for (int i = 0; i < statuses.Length; i++)
        {
            Cell cell = worksheet.Cells[i, 0];
            cell.PutValue(statuses[i]);
            ApplyStatusTheme(statuses[i], cell);
        }

        // Save the workbook
        workbook.Save("StatusThemeMapping.xlsx");
    }

    // Maps a business status to a specific theme accent color and applies it as cell fill
    static void ApplyStatusTheme(string status, Cell cell)
    {
        // Determine which ThemeColorType (accent) to use based on the status
        ThemeColorType accent;
        switch (status)
        {
            case "Open":
                accent = ThemeColorType.Accent1;
                break;
            case "Closed":
                accent = ThemeColorType.Accent2;
                break;
            case "Pending":
                accent = ThemeColorType.Accent3;
                break;
            case "InProgress":
                accent = ThemeColorType.Accent4;
                break;
            case "Cancelled":
                accent = ThemeColorType.Accent5;
                break;
            default:
                // Fallback accent for any undefined status
                accent = ThemeColorType.Accent6;
                break;
        }

        // Create a ThemeColor with no tint (0.0) for solid fill
        ThemeColor themeColor = new ThemeColor(accent, 0.0);

        // Retrieve the current style of the cell
        Style style = cell.GetStyle();

        // Apply the theme color as the foreground (fill) color and set a solid pattern
        style.ForegroundThemeColor = themeColor;
        style.Pattern = BackgroundType.Solid;

        // Assign the modified style back to the cell
        cell.SetStyle(style);
    }
}