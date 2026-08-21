// Title: Aspose.Cells for .NET: Create a Pie Chart with Percentage Labels and Auto‑Resizing Data‑Label Shapes
// Description: Shows how to build a workbook, fill category and value cells, add a pie chart, enable data labels to display only percentages, hide raw values, and automatically resize label shapes to fit the text using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# pie chart | percentage data labels | auto resize label shape | IsResizeShapeToFitText | WidthPixel | .NET charting | Excel export | pie chart formatting | data label customization
// Common Searches: Aspose.Cells show percentage on pie chart | C# auto resize chart data label shape | hide values display percentages Aspose.Cells chart | set minimum width for data labels Aspose.Cells | create pie chart from cells Aspose.Cells .NET
// Developer Intent: Generate a pie chart that reads values from worksheet cells, shows percentages in the data labels, hides the raw numbers, and lets the label shapes automatically adjust to the text length.
// Use Cases: Produce a sales‑by‑region report where each slice label shows only the percentage contribution and adapts its size to avoid truncation. | Build a dashboard that pulls percentage figures from cells, visualizes them in a pie chart, and maintains consistent label appearance across dynamic data sets. | Export a presentation slide with a pie chart whose data‑label boxes auto‑size, delivering a clean layout without manual adjustments.
// AI Prompts: Write C# code with Aspose.Cells to create a pie chart that displays percentages from cells and hides the raw values. | Show how to enable IsResizeShapeToFitText and define a minimum WidthPixel for data labels in an Aspose.Cells chart. | Explain the steps to bind category and value ranges to a pie chart and configure data‑label formatting, including percentage display and automatic shape resizing, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, fill category and value cells, add a pie chart, enable data labels to display only percentages, hide raw values, and automatically resize label shapes to fit the text using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and configure them to show percentages
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowPercentage = true;   // display percentage values
        dataLabels.ShowValue = false;       // hide raw values

        // Adjust the shape of the data label to fit the percentage text
        dataLabels.IsResizeShapeToFitText = true; // auto‑fit shape to text
        dataLabels.WidthPixel = 80;                // optional minimum width in pixels

        // Save the workbook to a file
        workbook.Save("PieChartWithPercentLabels.xlsx");
    }
}
