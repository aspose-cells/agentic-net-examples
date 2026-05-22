using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["A2"].PutValue("Product A");
        sheet.Cells["A3"].PutValue("Product B");
        sheet.Cells["A4"].PutValue("Product C");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and assign custom text to each point
        for (int s = 0; s < chart.NSeries.Count; s++)
        {
            Series series = chart.NSeries[s];
            series.DataLabels.ShowValue = true; // optional: show the value as part of the label

            for (int p = 0; p < series.Points.Count; p++)
            {
                ChartPoint point = series.Points[p];
                point.DataLabels.IsAutoText = false;               // turn off auto‑generated text
                point.DataLabels.Text = $"Sales: {point.YValue}";   // custom label for this point
            }
        }

        // Save the workbook with the customized data labels
        workbook.Save("CustomDataLabels.xlsx");
    }
}