// Title: Aspose.Cells .NET: Create a Combined Column‑Line Chart with Primary and Secondary Axes
// Description: Demonstrates how to build a mixed column and line chart in an Excel workbook using Aspose.Cells for .NET, assign the column series to the primary Y‑axis, move the line series to the secondary Y‑axis, and optionally label the secondary axis before saving the file.
// Keywords: Aspose.Cells combined chart .NET | column and line chart Aspose.Cells | secondary axis chart Aspose.Cells | mixed chart primary secondary axis | Aspose.Cells chart example C#
// Common Searches: Aspose.Cells create mixed column line chart | how to plot series on secondary axis Aspose.Cells | C# Aspose.Cells combined chart with two axes | set secondary Y axis title Aspose.Cells | Aspose.Cells chart type change line series
// Developer Intent: Generate a single Excel chart that shows column data on the primary Y‑axis and line data on a secondary Y‑axis using Aspose.Cells for .NET.
// Use Cases: Financial reports that compare monthly sales (columns) with profit margins (line). | Dashboard visualizations that display volume and percentage growth together. | Exporting analytical results to Excel with a mixed chart for stakeholder presentations.
// AI Prompts: Show how to change the line series color and marker style after moving it to the secondary axis in Aspose.Cells. | Provide code to add data labels to both column and line series in the combined chart. | Explain how to set custom number formats and define minimum/maximum values for the secondary Y‑axis.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to build a mixed column and line chart in an Excel workbook using Aspose.Cells for .NET, assign the column series to the primary Y‑axis, move the line series to the secondary Y‑axis, and optionally label the secondary axis before saving the file.
class CombinedColumnLineChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Populate sample data -----
        // Category (X) axis
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        // Column series (primary Y axis)
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Line series (secondary Y axis)
        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(60);

        // ----- Add a combined chart -----
        // Create a Column chart as the base chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add the column series (will use primary axis)
        chart.NSeries.Add("B2:B4", true);

        // Add the line series (will be moved to secondary axis)
        chart.NSeries.Add("C2:C4", true);

        // Set the category (X) axis data
        chart.NSeries.CategoryData = "A2:A4";

        // Configure the second series as a line and plot it on the secondary Y axis
        chart.NSeries[1].Type = ChartType.Line;          // change series type to line
        chart.NSeries[1].PlotOnSecondAxis = true;       // assign to secondary axis

        // Optional: give the secondary axis a title
        chart.SecondValueAxis.Title.Text = "Profit";

        // Save the workbook
        workbook.Save("CombinedColumnLineChart.xlsx");
    }
}
