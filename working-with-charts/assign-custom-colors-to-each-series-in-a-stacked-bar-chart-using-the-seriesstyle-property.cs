// Title: Set custom colors for each series in a horizontal stacked bar chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set series color in stacked bar chart | change individual series fill color in Aspose.Cells chart | apply custom solid colors to stacked bar series using Aspose.Cells .NET | how to use Series.Area.ForegroundColor to color chart series in C# | create horizontal stacked bar chart with custom series colors Aspose.Cells | Aspose.Cells chart series styling example C#
// Tags: Aspose.Cells series custom color | stacked bar chart series styling .NET | Series.Area.ForegroundColor usage | horizontal stacked bar chart Aspose.Cells | C# chart series fill color

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a horizontal stacked bar chart, and assigns distinct solid fill colors (Crimson, SeaGreen, SteelBlue) to each series by setting Series.Area.ForegroundColor and a transparent background, then saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the stacked bar chart
            // Categories (X axis)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Series 1
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(40);
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["B5"].PutValue(10);

            // Series 2
            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["C3"].PutValue(30);
            sheet.Cells["C4"].PutValue(40);
            sheet.Cells["C5"].PutValue(30);

            // Series 3
            sheet.Cells["D1"].PutValue("Product C");
            sheet.Cells["D2"].PutValue(10);
            sheet.Cells["D3"].PutValue(20);
            sheet.Cells["D4"].PutValue(30);
            sheet.Cells["D5"].PutValue(40);

            // Add a stacked bar chart (horizontal stacked bar)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart (including categories and series)
            chart.NSeries.Add("B2:D5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Define custom colors for each series
            Color[] seriesColors = new Color[] { Color.Crimson, Color.SeaGreen, Color.SteelBlue };

            // Apply custom colors to each series
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                // Set solid fill by specifying foreground color and transparent background
                series.Area.ForegroundColor = seriesColors[i];
                series.Area.BackgroundColor = Color.Transparent;
            }

            // Determine output file path
            string outputPath = "StackedBarCustomColors.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
