// Title: Thicken Bars in a 2‑D Bar (Progress) Chart by Setting GapWidth with Aspose.Cells for .NET
// Description: Shows how to build a workbook, add task‑progress data, create a 2‑D bar chart, and apply the Chart.GapWidth property (e.g., 30) to shrink the spacing between bars, making them appear thicker, then save the file as an XLSX document.
// Keywords: Aspose.Cells Chart.GapWidth | C# bar chart thickness | reduce gap width Aspose.Cells | increase bar width .NET | progress bar chart Excel | bar spacing 0‑500 | Excel chart formatting C# | thick bar chart Aspose | visualize progress with bars | gap width property example
// Common Searches: how to make bars thicker in Aspose.Cells bar chart | Aspose.Cells set GapWidth value | increase bar width for progress bar chart .NET | reduce spacing between bars Aspose.Cells | chart gap width range 0 to 500
// Developer Intent: The developer wants to decrease the chart's gap width so the bars look thicker and more visible.
// Use Cases: Design a dashboard that displays task progress with bold bar visuals. | Produce printed Excel reports where thicker bars improve readability. | Apply a uniform low GapWidth to multiple charts in a workbook for a consistent, bold appearance.
// AI Prompts: Provide C# code that sets Chart.GapWidth to 20 for all bar charts in a workbook using Aspose.Cells. | Explain the visual impact of GapWidth values from 0 to 500 on bar thickness and spacing. | Show how to export a workbook with thick‑bar progress charts to PDF after adjusting GapWidth.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, add task‑progress data, create a 2‑D bar chart, and apply the Chart.GapWidth property (e.g., 30) to shrink the spacing between bars, making them appear thicker, then save the file as an XLSX document.
class AdjustProgressBarGapWidth
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for a simple progress‑bar style chart
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["A2"].PutValue("Task 1");
        sheet.Cells["A3"].PutValue("Task 2");
        sheet.Cells["A4"].PutValue("Task 3");
        sheet.Cells["B1"].PutValue("Progress");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(60);
        sheet.Cells["B4"].PutValue(90);

        // Add a 2‑D bar chart (commonly used to represent progress bars)
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the data range to the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Decrease the gap width so the bars appear thicker (value between 0‑500)
        chart.GapWidth = 30;   // Smaller gap → thicker bars

        // Save the workbook
        workbook.Save("ProgressBarThick.xlsx", SaveFormat.Xlsx);
    }
}
