// Title: Apply a custom RGB color palette to each series of a column chart with Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a column chart using Aspose.Cells and assigns specific RGB colors to the fill of each series. | Show how to build a Color[] array and iterate over chart.NSeries to set Series.Area.ForegroundColor for custom series colors. | Explain how to programmatically customize the series colors of an Excel chart by setting the foreground color of each series area in Aspose.Cells.
// Common Searches: Aspose.Cells C# set individual series colors in a column chart using RGB values | custom color palette for Excel chart series with Aspose.Cells .NET | change fill color of each series in an Aspose.Cells generated chart programmatically
// Tags: set series foreground color Aspose.Cells | custom RGB palette Excel chart .NET | column chart series styling Aspose.Cells C# | apply custom colors to chart NSeries Aspose.Cells | Aspose.Cells chart series fill color

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The program creates a workbook, fills it with sample data, adds a column chart, defines an array of three RGB colors (red, green, blue), and assigns each color to the corresponding series' area foreground before saving the file as CustomPaletteChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sales Data";

            // Add series for each data column
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries.Add("D2:D4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define custom RGB colors for each series
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 0, 0),   // Red
                Color.FromArgb(0, 255, 0),   // Green
                Color.FromArgb(0, 0, 255)    // Blue
            };

            // Apply the custom colors to the series
            for (int i = 0; i < chart.NSeries.Count && i < customColors.Length; i++)
            {
                Series series = chart.NSeries[i];
                // Set custom fill color for the series
                series.Area.ForegroundColor = customColors[i];
            }

            // Save the workbook with the chart
            string outputPath = "CustomPaletteChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
