// Title: How to set a custom label on the first data point of each chart series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom DataLabel for first point in each chart series | Loop through chart series in Aspose.Cells and change point label text | Assign custom text to the first data point of a column chart with Aspose.Cells .NET | Programmatically style the first point label in an Aspose.Cells chart using C# | Enable and customize DataLabels for specific chart points in Aspose.Cells
// Tags: custom first-point DataLabel Aspose.Cells C# | chart series iteration Aspose.Cells .NET | modify point DataLabels text Aspose.Cells | column chart label styling Aspose.Cells | programmatic chart label assignment Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data and a column chart, then loops through each series to enable a DataLabel on the first point, sets custom text and styling, and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // First series values
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Second series values
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Series 1";

            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Series 2";

            // Set common category data for all series
            chart.NSeries.CategoryData = "A2:A4";

            // Loop through each series and assign a custom label to its first data point
            for (int s = 0; s < chart.NSeries.Count; s++)
            {
                Series series = chart.NSeries[s];

                // Ensure the series has at least one point
                if (series.Points.Count > 0)
                {
                    ChartPoint firstPoint = series.Points[0];

                    // Access the DataLabels object of the point
                    DataLabels pointLabel = firstPoint.DataLabels;

                    // Enable the label to be shown
                    pointLabel.ShowValue = true;

                    // Set a custom text for the label
                    pointLabel.Text = $"First of {series.Name}";

                    // Optional: style the label (font color, background, etc.)
                    pointLabel.Font.Color = Color.DarkBlue;
                    pointLabel.Font.IsBold = true;
                    pointLabel.BackgroundMode = BackgroundMode.Transparent;
                }
            }

            // Save the workbook to an XLSX file
            string outputPath = "CustomFirstPointLabel.xlsx";

            // Ensure the directory exists (CreateDirectory option can also be used)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
