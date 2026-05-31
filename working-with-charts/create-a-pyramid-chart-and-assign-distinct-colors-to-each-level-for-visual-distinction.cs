using Aspose.Cells;
using Aspose.Cells.Charts;

class PyramidChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pyramid chart
        sheet.Cells["A1"].PutValue("Level");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Level 1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Level 2");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Level 3");
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["A5"].PutValue("Level 4");
        sheet.Cells["B5"].PutValue(40);

        // Add a Pyramid chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pyramid, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries.CategoryData = "A2:A5";      // Categories

        // Enable varied colors so each level (data point) gets a distinct color
        chart.NSeries.IsColorVaried = true;

        // Optional: set a chart title
        chart.Title.Text = "Pyramid Chart with Distinct Colors";

        // Save the workbook
        workbook.Save("PyramidChartDistinctColors.xlsx");
    }
}