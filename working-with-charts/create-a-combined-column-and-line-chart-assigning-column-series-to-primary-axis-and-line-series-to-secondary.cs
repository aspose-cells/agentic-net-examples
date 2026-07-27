// Title: C# – Create a Combined Column & Line Chart with Primary & Secondary Axes using Aspose.Cells
// Description: This example shows how to build a mixed column‑line chart in a new workbook, assign the column series to the primary axis, switch the second series to a line chart, plot it on the secondary axis, set category data, add an optional secondary‑axis title, and save the result as CombinedColumnLineChart.xlsx.
// Keywords: Aspose.Cells combined chart | C# column line chart | secondary axis Aspose.Cells | mixed chart .NET | primary axis column series | line series secondary axis | Aspose.Cells chart example | Excel chart with two axes
// Common Searches: Aspose.Cells create mixed column and line chart | how to plot a series on secondary axis in Aspose.Cells | C# set secondary value axis for chart series | combined column‑line chart Aspose.Cells .NET | add line series to column chart Aspose.Cells
// Developer Intent: Generate a mixed column‑line chart where the column series uses the primary axis and the line series is displayed on a secondary axis.
// Use Cases: Display monthly sales as columns while showing profit trend as a line on a separate axis for financial dashboards. | Compare two metrics with different value ranges in a single chart for executive reports. | Create a presentation slide that highlights a primary KPI with columns and a secondary KPI with a line chart.
// AI Prompts: Show how to change the line series color and marker style in the combined chart. | Provide code to add data labels to both column and line series and format the secondary axis title. | Explain how to export the mixed chart as a PNG image using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to build a mixed column‑line chart in a new workbook, assign the column series to the primary axis, switch the second series to a line chart, plot it on the secondary axis, set category data, add an optional secondary‑axis title, and save the result as CombinedColumnLineChart.xlsx.
class CombinedColumnLineChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["A5"].PutValue("Apr");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(170);
        sheet.Cells["B5"].PutValue(200);

        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(50);
        sheet.Cells["C5"].PutValue(70);

        // Add a chart; start with a column chart type
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add the column series (Sales) and set its name
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries[0].Name = "Sales";

        // Add the line series (Profit) and set its name
        chart.NSeries.Add("C2:C5", true);
        chart.NSeries[1].Name = "Profit";

        // Change the second series to a line chart type
        chart.NSeries[1].Type = ChartType.Line;

        // Plot the line series on the secondary (right) value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Set the category (X) axis data
        chart.NSeries.CategoryData = "A2:A5";

        // Optional: give the secondary axis a title
        chart.SecondValueAxis.Title.Text = "Profit";

        // Save the workbook
        workbook.Save("CombinedColumnLineChart.xlsx");
    }
}
