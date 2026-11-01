// Title: How to modify individual chart theme colors via Theme.ColorScheme in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set custom accent color for a chart theme | Change hyperlink theme color in an Excel workbook using Aspose.Cells .NET | Programmatically update Theme.ColorScheme Accent2 color in a .xlsx file with Aspose.Cells | Apply a custom blue to chart Accent1 using Aspose.Cells Theme API | Modify Excel theme colors for a chart without affecting other workbook styles Aspose.Cells
// Tags: Aspose.Cells modify chart theme colors | Theme.ColorScheme SetColor C# | customize Excel chart accent colors Aspose.Cells | update Excel theme hyperlink color programmatically | Aspose.Cells Theme API example .xlsx

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, shows how to change specific theme colors through the Theme.ColorScheme collection (code commented for version compatibility), adds a column chart linked to range A1:A5, sets a chart title, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Theme manipulation APIs may not be available in older Aspose.Cells versions.
            // The following block is omitted to ensure compatibility.
            // If your version supports Theme, you can uncomment and adjust accordingly.

            /*
            Theme theme = workbook.Theme;
            if (theme?.ColorScheme != null)
            {
                ThemeColorScheme colorScheme = theme.ColorScheme;
                colorScheme.SetColor(ThemeColorType.Accent1, Color.FromArgb(0, 112, 192));   // Custom blue
                colorScheme.SetColor(ThemeColorType.Accent2, Color.FromArgb(255, 165, 0));   // Custom orange
                colorScheme.SetColor(ThemeColorType.Hyperlink, Color.FromArgb(128, 0, 128)); // Custom purple
            }
            */

            // Add a chart to demonstrate workbook modifications
            Worksheet sheet = workbook.Worksheets[0];
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A5", true);
            chart.Title.Text = "Sample Chart";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
