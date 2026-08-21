// Title: Aspose.Cells for .NET – Configure Chart Axes and Freeze the Axis‑Label Row
// Description: Creates a workbook, adds sample data, inserts a column chart, customizes X and Y axes (titles, tick marks, fixed min/max values), forces chart calculation to generate axis labels, then freezes the first worksheet row that contains the category axis labels before saving the file.
// Keywords: Aspose.Cells chart axis customization | C# set chart axis min max | Aspose.Cells FreezePanes row | calculate chart before freezing Aspose.Cells | column chart Aspose.Cells .NET | Excel FreezePanes C# example | chart axis titles Aspose.Cells | tick marks chart Aspose.Cells
// Common Searches: how to freeze row with chart axis labels using Aspose.Cells | set custom minimum and maximum values for a chart axis in C# | add titles and tick marks to Aspose.Cells chart axes | Aspose.Cells FreezePanes example for Excel worksheets | calculate chart before applying FreezePanes Aspose.Cells
// Developer Intent: Add a column chart, tailor its axes, and lock the worksheet row that holds the category labels.
// Use Cases: Generate a sales column chart with defined axis titles and a fixed scale, then keep the quarter header visible while scrolling through extensive data. | Produce a financial dashboard where major and minor tick marks are specified and the top row is frozen to maintain axis label context. | Create an automated Excel report that calculates the chart to render axis labels before applying FreezePanes, ensuring static labels during navigation.
// AI Prompts: Provide C# code using Aspose.Cells to set custom min/max values for a chart's value axis and freeze the first row. | Show an example that adds titles to both category and value axes of a column chart and then freezes the row containing the category labels. | Explain why chart.Calculate() must be called before FreezePanes when working with Aspose.Cells charts, and give the complete workflow. | Write a snippet that configures tick marks for chart axes and applies FreezePanes to keep axis labels visible in a large worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, customizes X and Y axes (titles, tick marks, fixed min/max values), forces chart calculation to generate axis labels, then freezes the first worksheet row that contains the category axis labels before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];

        // Populate sample data: categories in column A, values in column B
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("Q1");
        ws.Cells["A3"].PutValue("Q2");
        ws.Cells["A4"].PutValue("Q3");
        ws.Cells["B2"].PutValue(120);
        ws.Cells["B3"].PutValue(150);
        ws.Cells["B4"].PutValue(180);

        // Add a column chart to the worksheet
        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = ws.Charts[chartIdx];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure the Category (X) axis
        chart.CategoryAxis.Title.Text = "Quarter";
        chart.CategoryAxis.Title.IsVisible = true;
        chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;
        chart.CategoryAxis.MinorTickMark = TickMarkType.Inside;

        // Configure the Value (Y) axis
        chart.ValueAxis.Title.Text = "Sales";
        chart.ValueAxis.Title.IsVisible = true;
        chart.ValueAxis.IsAutomaticMinValue = false;
        chart.ValueAxis.MinValue = 0;
        chart.ValueAxis.IsAutomaticMaxValue = false;
        chart.ValueAxis.MaxValue = 200;
        chart.ValueAxis.MajorUnit = 50;
        chart.ValueAxis.MinorUnit = 10;

        // Calculate the chart to ensure axis labels are generated
        chart.Calculate();

        // Freeze the row that contains the axis (category) labels (row 1)
        // FreezePanes(row, column, freezedRows, freezedColumns)
        // Row and column are 1‑based indexes where the split occurs.
        ws.FreezePanes(2, 1, 1, 0); // Freeze first row, no frozen columns

        // Save the workbook
        workbook.Save("ChartWithFrozenAxisLabels.xlsx");
    }
}
