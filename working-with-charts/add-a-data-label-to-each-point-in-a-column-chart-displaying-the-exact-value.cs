using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

class AddDataLabelsToColumnChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Get the first series (there is only one)
            Series series = chart.NSeries[0];

            // Enable data labels for the series and set a position (optional)
            series.DataLabels.ShowValue = true;               // show the value automatically
            series.DataLabels.Position = LabelPositionType.OutsideEnd;

            // Set a custom label for each point to display the exact value
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                if (point != null && point.DataLabels != null)
                {
                    // Use the YValue of the point as the label text
                    point.DataLabels.Text = point.YValue.ToString();
                }
            }

            // Save the workbook with the chart and data labels
            workbook.Save("ColumnChartWithDataLabels.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}