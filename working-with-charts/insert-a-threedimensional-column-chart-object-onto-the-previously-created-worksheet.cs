using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["B4"].PutValue(1800);

        // Add a 3‑D column chart using the ChartCollection.Add method
        int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Optional: adjust 3‑D visual properties
        chart.Elevation = 30;          // Elevation angle in degrees
        chart.RotationAngle = 45;      // Rotation angle in degrees
        chart.DepthPercent = 150;      // Depth as a percentage of chart width

        // Save the workbook with the chart
        workbook.Save("ThreeDColumnChart.xlsx", SaveFormat.Xlsx);
    }
}