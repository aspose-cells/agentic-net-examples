// Title: Create an Excel workbook that lists all Aspose.Cells ThemeColorType names with their #RRGGBB hex values in C#
// AI Prompts: Write C# code using Aspose.Cells to loop through the ThemeColorType enum, retrieve each color with GetThemeColor, format it as a #RRGGBB string, and place the name and hex code into two columns of a new worksheet. | Add header cells "Theme Color Name" and "Hex Value" to the worksheet, auto‑fit the columns, and save the file as ThemeSummary.xlsx. | Implement a method that returns a DataTable containing ThemeColorType names paired with their hexadecimal values, then export that table to an .xlsx file with Aspose.Cells.
// Common Searches: C# Aspose.Cells how to list all theme colors with hex codes | GetThemeColor enum values to Excel using Aspose.Cells | Export Aspose.Cells theme palette to .xlsx file in C# | Convert Aspose.Cells theme colors to #RRGGBB strings programmatically | Auto fit columns after writing data with Aspose.Cells C#
// Tags: Aspose.Cells enumerate ThemeColorType to Excel | C# extract theme colors hex values | write theme color summary worksheet | auto‑fit columns Aspose.Cells workbook | save theme palette as xlsx file

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a worksheet named "Theme Summary", writes headers, iterates over every ThemeColorType enum value, obtains each theme color via GetThemeColor, converts it to a #RRGGBB string, records the name and hex value in two columns, auto‑fits the columns, and saves the result as ThemeSummary.xlsx.
class ThemeColorSummary
{
    static void Main()
    {
        try
        {
            // Create a new workbook (uses the default theme)
            Workbook workbook = new Workbook();

            // Add a worksheet to hold the theme color summary
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[sheetIndex];
            summarySheet.Name = "Theme Summary";

            // Write header titles
            summarySheet.Cells[0, 0].PutValue("Theme Color Name");
            summarySheet.Cells[0, 1].PutValue("Hex Value");

            // Iterate through all ThemeColorType values and retrieve their colors
            int row = 1;
            foreach (ThemeColorType themeType in Enum.GetValues(typeof(ThemeColorType)))
            {
                Color color = workbook.GetThemeColor(themeType);
                string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";

                summarySheet.Cells[row, 0].PutValue(themeType.ToString());
                summarySheet.Cells[row, 1].PutValue(hex);
                row++;
            }

            // Auto‑fit columns for better readability
            summarySheet.AutoFitColumns();

            // Save the workbook
            string outputPath = "ThemeSummary.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
