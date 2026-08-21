// Title: Create a Mixed Column‑Line Chart with Primary and Secondary Axes Using Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to build an Excel workbook with month, sales, and profit data, add a column chart, assign the sales series to the primary axis, add a profit series as a line plotted on the secondary axis, set axis titles, and save the file as CombinedColumnLineChart.xlsx.
// Keywords: Aspose.Cells | C# chart example | combined column line chart | secondary axis Aspose.Cells | mixed chart .NET | plot line on secondary axis | Excel chart Aspose | ChartType.Column | ChartType.Line
// Common Searches: Aspose.Cells add line series to column chart | C# mixed column and line chart Aspose | set secondary Y axis in Aspose.Cells | create combined chart with primary and secondary axes .NET | Aspose.Cells chart example secondary axis
// Developer Intent: Create a mixed column‑line chart where the column series uses the primary Y‑axis and the line series uses a secondary Y‑axis.
// Use Cases: Show monthly sales as columns and profit as a line with separate scales for clear comparison. | Compare metrics with different units (e.g., units sold vs. profit margin) in a single visualization. | Design dashboard widgets that combine absolute values and trend lines. | Prepare financial reports where volume and rate need distinct axes.
// AI Prompts: Generate C# code with Aspose.Cells to build a combined column and line chart, placing the line series on a secondary axis and adding axis titles. | Explain how to add an additional stacked column series to the primary axis in the mixed chart example. | Provide step‑by‑step documentation for configuring primary and secondary axes in Aspose.Cells charts. | Suggest ways to customize colors and markers for the line series in the combined chart.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example demonstrates how to build an Excel workbook with month, sales, and profit data, add a column chart, assign the sales series to the primary axis, add a profit series as a line plotted on the secondary axis, set axis titles, and save the file as CombinedColumnLineChart.xlsx.
class CombinedColumnLineChartDemo
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

        sheet.Cells["B1"].PutValue("Sales");          // Column series data
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        sheet.Cells["C1"].PutValue("Profit");         // Line series data
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(60);

        // Add a chart (initially a Column chart)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add the column series (primary Y axis)
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries[0].Type = ChartType.Column; // explicit, though default

        // Add the line series (secondary Y axis)
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries[1].Type = ChartType.Line;
        chart.NSeries[1].PlotOnSecondAxis = true; // place on secondary axis

        // Set category (X) axis data
        chart.NSeries.CategoryData = "A2:A4";

        // Make secondary value axis visible and give it a title
        chart.SecondValueAxis.IsVisible = true;
        chart.SecondValueAxis.Title.Text = "Profit";

        // Save the workbook
        workbook.Save("CombinedColumnLineChart.xlsx");
    }
}
