// Title: C# – Set a Custom Data Label on the First Point of Each Chart Series with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, adds two data series, builds a column chart, and loops through every series to enable a data label only on the first point. The label text is set to a custom string that includes the series number, the chart is recalculated, and the file is saved as an Excel workbook.
// Keywords: Aspose.Cells | .NET | C# | chart series first point label | custom data label | ChartPoint DataLabels | loop through chart series | column chart example | Excel automation | GitHub sample
// Common Searches: Aspose.Cells set custom label on first chart point | C# loop chart series to change data label | how to show label only for first point in Aspose.Cells chart | assign text to first data point in Excel chart using Aspose | Aspose.Cells example for custom data labels
// Developer Intent: Apply a unique text label to the first data point of every series in an Excel chart generated with Aspose.Cells.
// Use Cases: Highlight the initial value of each series with a distinct label. | Add a series identifier (e.g., "Series 1 – First") only to the first column of a chart. | Create cleaner charts by displaying data labels selectively on the first point of each series.
// AI Prompts: Write C# code using Aspose.Cells that iterates over all chart series and sets a custom data label on the first point of each series. | Show how to hide data labels for all points except the first one in each series and customize the label text in an Aspose.Cells column chart. | Provide an Aspose.Cells example that changes the font style and color of a custom label applied to the first data point of each series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomFirstPointLabel
{
    // This Aspose.Cells for .NET example creates a workbook, adds two data series, builds a column chart, and loops through every series to enable a data label only on the first point. The label text is set to a custom string that includes the series number, the chart is recalculated, and the file is saved as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Loop through each series and assign a custom label to its first data point
            for (int s = 0; s < chart.NSeries.Count; s++)
            {
                Series series = chart.NSeries[s];

                // Ensure the series has at least one point
                if (series.Points.Count > 0)
                {
                    // Access the first point (index 0)
                    ChartPoint firstPoint = series.Points[0];

                    // Enable the data label for this point
                    firstPoint.DataLabels.ShowValue = true;

                    // Assign a custom text to the data label
                    firstPoint.DataLabels.Text = $"Series {s + 1} - First";
                }
            }

            // Recalculate the chart to apply changes
            chart.Calculate();

            // Save the workbook
            workbook.Save("CustomFirstPointLabels.xlsx");
        }
    }
}
