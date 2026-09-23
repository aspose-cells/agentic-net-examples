// Title: Create a C# batch program to apply a spring seasonal theme to all quarterly report Excel workbooks using Aspose.Cells
// AI Prompts: Write a C# console application that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, and applies predefined header and data cell styles for a spring theme before saving. | Develop a reusable method that defines seasonal colors, creates header and data styles, applies them to every worksheet in an Aspose.Cells workbook, auto‑fits columns, and returns the styled workbook.
// Common Searches: C# Aspose.Cells batch apply header background color to multiple Excel files | How to style all worksheets in a folder of .xlsx files with Aspose.Cells | Apply seasonal color scheme to quarterly report workbooks using Aspose.Cells | Automate Excel workbook theming across many files in .NET | Set background and font colors for header rows in all sheets with Aspose.Cells
// Tags: batch apply seasonal style Aspose.Cells | format header cells Excel C# | apply background color to data cells Aspose.Cells | auto-fit columns after styling Aspose.Cells | process multiple .xlsx workbooks C#

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// A C# console utility that iterates through all .xlsx files in a specified folder, loads each workbook with Aspose.Cells, applies a spring‑themed header and data cell style to every worksheet, auto‑fits columns, and saves the themed workbooks to an output directory.
class SeasonalThemeBatch
{
    // Define the seasonal colors (example: Spring theme)
    private static readonly Color HeaderBackground = Color.LightGreen;
    private static readonly Color DataBackground = Color.Honeydew;
    private static readonly Color HeaderFontColor = Color.DarkGreen;
    private static readonly Color DataFontColor = Color.Black;

    static void Main(string[] args)
    {
        // Folder containing quarterly report workbooks
        string sourceFolder = @"C:\QuarterlyReports";
        // Folder to save themed workbooks (can be same as source)
        string outputFolder = @"C:\QuarterlyReports\Themed";

        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            // Load workbook
            Workbook workbook = new Workbook(filePath);

            // Apply seasonal theme
            ApplySeasonalTheme(workbook);

            // Build output file path
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save workbook
            workbook.Save(outputPath);
        }

        Console.WriteLine("Seasonal theme applied to all quarterly reports.");
    }

    private static void ApplySeasonalTheme(Workbook workbook)
    {
        // Create a style for header cells
        Style headerStyle = workbook.CreateStyle();
        headerStyle.ForegroundColor = HeaderBackground;
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.Font.Color = HeaderFontColor;
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Size = 12;

        // Create a style for data cells
        Style dataStyle = workbook.CreateStyle();
        dataStyle.ForegroundColor = DataBackground;
        dataStyle.Pattern = BackgroundType.Solid;
        dataStyle.Font.Color = DataFontColor;
        dataStyle.Font.Size = 11;

        // Apply styles to each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Assume first row is header
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Apply header style to first row
            for (int col = 0; col <= maxCol; col++)
            {
                Cell headerCell = cells[0, col];
                headerCell.SetStyle(headerStyle);
            }

            // Apply data style to remaining rows
            for (int row = 1; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell dataCell = cells[row, col];
                    dataCell.SetStyle(dataStyle);
                }
            }

            // Optionally adjust column widths for better appearance
            sheet.AutoFitColumns();
        }
    }
}
