// Title: How to Thicken Bars in an Aspose.Cells Bar Chart by Reducing GapWidth (C#)
// Description: Demonstrates creating a workbook, adding task‑progress data, inserting a Bar chart, and setting the Chart.GapWidth property to a low percentage (e.g., 30) so the bars appear thicker and more visible, then saving the file as XLSX.
// Keywords: Aspose.Cells GapWidth | C# bar chart thickness | increase bar width Aspose.Cells | progress bar chart Aspose.Cells | reduce chart gap width | thick bar chart Excel | Aspose.Cells chart spacing
// Common Searches: Aspose.Cells set GapWidth for bar chart | make bar chart columns thicker C# | adjust bar spacing Aspose.Cells | progress bar style chart Aspose.Cells | increase bar thickness in Excel using Aspose
// Developer Intent: Lower the chart's GapWidth value to produce thicker bars.
// Use Cases: Designing progress‑bar visualizations where bar width must stand out on dashboards. | Creating printable reports with bold bar charts for improved readability. | Optimizing clustered bar charts for slide presentations by minimizing gaps.
// AI Prompts: Generate C# code that creates a clustered column chart with Aspose.Cells and sets GapWidth to 20, explaining the visual impact. | Write a function that calculates an optimal GapWidth based on the number of series in a bar chart using Aspose.Cells. | Provide a step‑by‑step guide to build a progress‑bar chart in Aspose.Cells, configure GapWidth for thick bars, and export to XLSX.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, adding task‑progress data, inserting a Bar chart, and setting the Chart.GapWidth property to a low percentage (e.g., 30) so the bars appear thicker and more visible, then saving the file as XLSX.
class AdjustProgressBarGapWidth
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a progress‑bar style chart
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["A2"].PutValue("Task 1");
        sheet.Cells["A3"].PutValue("Task 2");
        sheet.Cells["A4"].PutValue("Task 3");

        sheet.Cells["B1"].PutValue("Progress");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(70);
        sheet.Cells["B4"].PutValue(55);

        // Add a bar chart (commonly used to represent progress bars)
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Decrease the gap width to make the bars appear thicker
        // GapWidth is a percentage of the bar width (0‑500). Lower values = thicker bars.
        chart.GapWidth = 30;

        // Save the workbook with the modified chart
        workbook.Save("ProgressBarGapWidth.xlsx", SaveFormat.Xlsx);
    }
}
