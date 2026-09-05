// Title: Insert a three‑dimensional column chart into an Aspose.Cells worksheet using C#
// AI Prompts: Write C# code that creates a workbook, populates cells A1:B4 with data, inserts a 3‑D column chart positioned from rows 5 to 20 and columns A to I, binds the chart to the data range, verifies the chart’s Is3D flag, and exports the file as an XLSX workbook. | Show how to employ Aspose.Cells ChartType.Column3D to generate a three‑dimensional column chart, assign its source range, and save the workbook containing the chart.
// Common Searches: Aspose.Cells C# example for creating a three-dimensional column chart and exporting to XLSX | how to bind a three-dimensional column chart to a specific data range using Aspose.Cells | setting chart position and checking 3D flag in Aspose.Cells .NET
// Tags: Aspose.Cells add three‑dimensional column chart via C# | Aspose.Cells Column3D chart type usage | Aspose.Cells define chart data source | Aspose.Cells determine if chart is 3D | Aspose.Cells export workbook containing chart to XLSX

using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes sample sales data into cells A1:B4, inserts a three‑dimensional column chart (ChartType.Column3D) positioned between rows 5‑20 and columns A‑I, sets the chart’s source range to A1:B4, checks the Is3D flag to confirm the chart is 3D, and saves the workbook as ThreeDColumnChart.xlsx.
class Insert3DColumnChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["B4"].PutValue(1800);

        // Add a three‑dimensional column chart to the worksheet
        // Parameters: chart type, top row, left column, bottom row, right column
        int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (including headers)
        chart.SetChartDataRange("A1:B4", true);

        // Verify that the chart is indeed 3D
        bool is3D = chart.Is3D; // Expected to be true

        // Save the workbook with the newly added chart
        workbook.Save("ThreeDColumnChart.xlsx", SaveFormat.Xlsx);
    }
}
