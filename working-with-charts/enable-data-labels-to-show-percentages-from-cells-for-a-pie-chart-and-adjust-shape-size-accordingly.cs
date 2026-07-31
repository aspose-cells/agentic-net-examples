// Title: Aspose.Cells for .NET – Display Cell‑Based Percentages on a Pie Chart and Auto‑Fit Label Shapes
// Description: This C# sample builds a workbook, fills it with categories and values, inserts a pie chart, configures the series, turns on data labels to show percentages while hiding raw numbers, and automatically resizes each label shape to fit its text with a preset minimum width before saving the file.
// Keywords: Aspose.Cells | .NET | C# | pie chart | data labels | percentage label | auto‑fit label shape | IsResizeShapeToFitText | WidthPixel | Excel export
// Common Searches: Aspose.Cells show percentage on pie chart | auto resize data label shape Aspose.Cells | C# pie chart label width | display only percentages in Excel chart using Aspose | adjust pie chart data label size programmatically
// Developer Intent: The developer wants to present percentage values taken from worksheet cells on a pie‑chart’s data labels and have the label containers automatically adjust their dimensions to the text.
// Use Cases: Generate a sales‑distribution pie chart that displays only percentages and ensures each label expands to fit its content. | Create a marketing report where pie‑chart labels hide numeric values, show calculated percentages, and maintain a readable minimum width. | Automate Excel exports that format pie‑chart data labels for dynamic data sets, preserving visual consistency across different volumes.
// AI Prompts: Write C# code with Aspose.Cells to add a pie chart, show only percentages from cells in the data labels, and set the label shape to auto‑fit with a minimum width of 80 pixels. | Explain the impact of DataLabels.IsResizeShapeToFitText and DataLabels.WidthPixel on pie‑chart label rendering in Aspose.Cells. | Provide a step‑by‑step tutorial for customizing pie‑chart data labels to display percentages and automatically adjust label size in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPieChartDemo
{
    // This C# sample builds a workbook, fills it with categories and values, inserts a pie chart, configures the series, turns on data labels to show percentages while hiding raw numbers, and automatically resizes each label shape to fit its text with a preset minimum width before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(20);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            pieChart.NSeries.Add("B2:B4", true);
            pieChart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and show percentages from the cells
            DataLabels dataLabels = pieChart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;   // display percentage values
            dataLabels.ShowValue = false;       // hide raw values (optional)

            // Adjust the shape of the data labels to fit the text automatically
            dataLabels.IsResizeShapeToFitText = true;   // auto‑fit shape size
            // Optionally set a minimum size; here we set a reasonable width in pixels
            dataLabels.WidthPixel = 80;

            // Save the workbook to a file
            workbook.Save("PieChartWithPercentLabels.xlsx");
        }
    }
}
