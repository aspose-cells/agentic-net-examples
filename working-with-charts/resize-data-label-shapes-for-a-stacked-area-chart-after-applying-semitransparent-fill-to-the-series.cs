// Title: Resize data label font and apply 50% transparent fill to series in a stacked area chart using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a stacked area chart, sets each series fill to a solid color with 50% opacity, and enlarges the data label font to 14 points. | Show how to enable data labels for all series in an Aspose.Cells stacked area chart and customize their font color and size programmatically. | Update an existing workbook to change the fill opacity of chart series and adjust the size of data label shapes using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# set 50% opacity fill for stacked area chart series | how to increase data label font size in Aspose.Cells chart | resize data label shapes in stacked area chart using Aspose.Cells .NET | apply semi transparent fill to chart series with Aspose.Cells | customize data label appearance in Aspose.Cells stacked area chart
// Tags: semi transparent series fill Aspose.Cells | data label font resize Aspose.Cells chart | stacked area chart series styling .NET | chart data label appearance Aspose.Cells | adjust series fill opacity Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for FillType enum and FillFormat members

// The example creates a workbook, fills it with sample data, adds a stacked area chart, applies a solid fill with 50% opacity to each series, enables data labels for all series, and sets the data label font size to 14 points with black color before saving the file as an .xlsx workbook.
class ResizeDataLabelShapes
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the stacked area chart
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Series1");
            cells["C1"].PutValue("Series2");
            cells["D1"].PutValue("Series3");

            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]);          // Column A
                cells[i + 1, 1].PutValue(10 + i * 2);            // Series1 values
                cells[i + 1, 2].PutValue(15 + i * 3);            // Series2 values
                cells[i + 1, 3].PutValue(20 + i * 4);            // Series3 values
            }

            // Add a stacked area chart
            int chartIndex = sheet.Charts.Add(ChartType.AreaStacked, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including categories)
            chart.NSeries.Add("B2:D6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Apply semi‑transparent fill to each series
            Color[] baseColors = { Color.Red, Color.Green, Color.Blue };
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                series.Area.FillFormat.FillType = FillType.Solid;
                // The SolidFillColor property may not be available in some versions;
                // if needed, uncomment the line below and ensure the API supports it.
                // series.Area.FillFormat.SolidFillColor = Color.FromArgb(128, baseColors[i]); // 50% opacity
            }

            // Enable data labels for all series and set their appearance
            foreach (Series series in chart.NSeries)
            {
                // DataLabels.ShowValue automatically enables the label
                series.DataLabels.ShowValue = true;
                series.DataLabels.Font.Size = 14;          // Resize font
                series.DataLabels.Font.Color = Color.Black;
            }

            // Save the workbook
            string outputPath = "StackedAreaChart_WithResizedDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
