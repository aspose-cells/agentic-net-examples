// Title: Set Data Labels Inside End of Columns in Aspose.Cells C# Column Chart
// Description: Step‑by‑step guide to create a column chart with Aspose.Cells, enable data labels and position them inside the column ends to avoid overlapping the bars, using C#.
// Keywords: Aspose.Cells C# column chart | data label position InsideEnd | prevent overlapping chart labels | Aspose.Cells chart data labels | LabelPositionType InsideEnd example | Excel column chart Aspose.Cells | C# Aspose.Cells chart tutorial
// Common Searches: Aspose.Cells set data label inside end column chart | C# place column chart labels at top inside bar | avoid overlapping data labels Aspose.Cells | how to use LabelPositionType.InsideEnd in Aspose.Cells | Aspose.Cells column chart label positioning
// Developer Intent: Position data labels at the inside end of each column so the values appear inside the bar tops without covering the columns.
// Use Cases: Generate Excel reports where column values are displayed inside the bar for a cleaner look. | Automate workbook creation with Aspose.Cells while ensuring chart labels never overlap the graphics. | Build reusable chart templates that consistently apply the InsideEnd label style for better readability.
// AI Prompts: Show C# code to set data label position to InsideEnd for an Aspose.Cells column chart. | Explain how LabelPositionType.InsideEnd affects the appearance of column chart labels in Aspose.Cells. | Provide a sample that creates a column chart with data labels positioned inside the column ends to avoid overlap.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Step‑by‑step guide to create a column chart with Aspose.Cells, enable data labels and position them inside the column ends to avoid overlapping the bars, using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Fill sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and position them inside the end of each column
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;                     // Show the numeric value
        dataLabels.Position = LabelPositionType.InsideEnd; // Position inside the end of the column

        // Save the workbook
        workbook.Save("ColumnChart_InsideEnd.xlsx");
    }
}
