// Title: Set Bar Chart GapWidth in Aspose.Cells for .NET to Thicken Progress Bars
// Description: Creates a workbook, adds task‑progress data, inserts a 2‑D bar chart, sets GapWidth to 30% to reduce spacing between bars, and saves the file as ProgressBarThick.xlsx.
// Keywords: Aspose.Cells | C# | BarChart | GapWidth | thick bars | progress bar chart | chart styling | Excel automation
// Common Searches: Aspose.Cells set GapWidth | make bar chart bars thicker Aspose.Cells | reduce gap width progress bar Aspose.Cells .NET | increase bar thickness Excel chart programmatically | adjust bar spacing Aspose.Cells
// Developer Intent: Decrease the GapWidth of a bar chart so the bars appear thicker, improving visibility of progress‑bar style charts.
// Use Cases: Design project‑status dashboards with bold progress bars. | Generate Excel reports where bar charts need stronger visual emphasis. | Batch‑process worksheets to standardize bar thickness across multiple charts. | Create printable charts for presentations with minimal gaps.
// AI Prompts: Write C# code using Aspose.Cells to set GapWidth to 20 for a column chart and explain the visual impact. | Show how to calculate an optimal GapWidth based on the number of series in a bar chart. | Create a sample that iterates through all charts in a workbook and sets GapWidth to 40. | Explain the difference between GapWidth and Overlap properties in Aspose.Cells charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds task‑progress data, inserts a 2‑D bar chart, sets GapWidth to 30% to reduce spacing between bars, and saves the file as ProgressBarThick.xlsx.
class AdjustProgressBarGapWidth
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a progress‑bar style chart
        worksheet.Cells["A1"].PutValue("Task");
        worksheet.Cells["A2"].PutValue("Design");
        worksheet.Cells["A3"].PutValue("Development");
        worksheet.Cells["A4"].PutValue("Testing");

        worksheet.Cells["B1"].PutValue("Progress");
        worksheet.Cells["B2"].PutValue(70);
        worksheet.Cells["B3"].PutValue(40);
        worksheet.Cells["B4"].PutValue(90);

        // Add a 2‑D bar chart (commonly used for progress bars)
        int chartIndex = worksheet.Charts.Add(ChartType.Bar, 6, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Decrease the gap width to make the bars appear thicker.
        // GapWidth is a percentage of the bar width (0‑500). Lower values = less space.
        chart.GapWidth = 30;   // 30% of the default gap, resulting in thicker bars

        // Save the workbook with the adjusted chart
        workbook.Save("ProgressBarThick.xlsx", SaveFormat.Xlsx);
    }
}
