using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (two series)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIdx];

        // Add the two series (vertical orientation)
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Loop through each series and assign a custom label to its first data point
        for (int s = 0; s < chart.NSeries.Count; s++)
        {
            Series series = chart.NSeries[s];

            // Ensure the series contains at least one point
            if (series.Points.Count > 0)
            {
                // Access the first point (index 0)
                ChartPoint firstPoint = series.Points[0];

                // Enable the data label for this point (optional, but ensures visibility)
                firstPoint.DataLabels.ShowValue = true;

                // Set a custom text for the data label
                firstPoint.DataLabels.Text = $"First point of series {s + 1}";

                // Optionally, set the position of the label
                firstPoint.DataLabels.Position = LabelPositionType.Center;
            }
        }

        // Recalculate the chart to apply changes
        chart.Calculate();

        // Save the workbook
        workbook.Save("CustomFirstPointLabels.xlsx");
    }
}