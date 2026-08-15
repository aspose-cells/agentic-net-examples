// Title: C# – Map Business Status to Excel Theme Accent Colors Using Aspose.Cells
// Description: Demonstrates how to create a function that returns a ThemeColor (accent type with optional tint) for a given business status and apply it as a solid cell background. The sample writes status values, sets the ForegroundThemeColor, and saves the workbook as StatusThemeMapping.xlsx.
// Keywords: Aspose.Cells C# | ThemeColor mapping | Excel theme accent fill | status based cell color | C# cell style background | ThemeColorType Accent1 Accent2 | Excel workbook generation .NET | conditional formatting Aspose
// Common Searches: Aspose.Cells map status to theme color C# | set cell background using ThemeColor in .NET | apply Excel theme accent colors with Aspose | C# function to color‑code status cells | how to use ThemeColorType in Aspose.Cells
// Developer Intent: Create a reusable method that translates a business status string into an Aspose.Cells ThemeColor and use it to color‑code cell fills.
// Use Cases: Color‑code task‑list worksheets so each status (Open, In Progress, Closed, etc.) displays a distinct theme accent. | Build a status dashboard where rows are visually differentiated by corporate theme colors. | Ensure consistent branding across multiple workbooks by mapping status values to predefined theme accents.
// AI Prompts: Generate a C# method that receives a status string and returns an Aspose.Cells ThemeColor with the appropriate Accent and tint. | Show how to assign the returned ThemeColor to the ForegroundThemeColor of a cell style for a range of cells. | Adapt the status‑to‑ThemeColor mapping to use custom RGB colors instead of built‑in theme accents.

using System;
using Aspose.Cells;

namespace AsposeCellsStatusThemeDemo
{
    // Demonstrates how to create a function that returns a ThemeColor (accent type with optional tint) for a given business status and apply it as a solid cell background. The sample writes status values, sets the ForegroundThemeColor, and saves the workbook as StatusThemeMapping.xlsx.
    public class StatusThemeMapper
    {
        // Maps a business status string to a ThemeColor (accent color with optional tint)
        private static ThemeColor GetThemeColorForStatus(string status)
        {
            // Define mapping between status values and theme accent types
            // Adjust tint values as needed (0 = no tint, positive = lighten, negative = darken)
            switch (status.Trim().ToLower())
            {
                case "new":
                case "open":
                    return new ThemeColor(ThemeColorType.Accent1, 0.0);   // Accent1
                case "in progress":
                case "pending":
                    return new ThemeColor(ThemeColorType.Accent2, 0.2);   // Accent2, slightly lighter
                case "completed":
                case "closed":
                    return new ThemeColor(ThemeColorType.Accent3, -0.2);  // Accent3, slightly darker
                case "on hold":
                    return new ThemeColor(ThemeColorType.Accent4, 0.0);   // Accent4
                case "canceled":
                    return new ThemeColor(ThemeColorType.Accent5, -0.4);  // Accent5, darker
                default:
                    // Fallback to a neutral accent
                    return new ThemeColor(ThemeColorType.Accent6, 0.0);   // Accent6
            }
        }

        // Demonstrates applying the theme colors to cells based on status values
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data: status values in column A
                string[] statuses = { "Open", "Pending", "Closed", "On Hold", "Canceled", "Unknown" };
                for (int i = 0; i < statuses.Length; i++)
                {
                    // Write status text
                    Cell cell = cells[i, 0]; // Column A
                    cell.PutValue(statuses[i]);

                    // Create a style for the cell
                    Style style = workbook.CreateStyle();
                    style.Pattern = BackgroundType.Solid;

                    // Retrieve the theme color for the current status
                    ThemeColor themeColor = GetThemeColorForStatus(statuses[i]);

                    // Apply the theme color as the cell's background (foreground theme color)
                    style.ForegroundThemeColor = themeColor;

                    // Assign the style to the cell
                    cell.SetStyle(style);
                }

                // Save the workbook (lifecycle rule: save)
                workbook.Save("StatusThemeMapping.xlsx");
                Console.WriteLine("Workbook saved as 'StatusThemeMapping.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while generating the workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            StatusThemeMapper.Run();
        }
    }
}
