// Title: Aspose.Cells C# Example: Create a Stacked Column Chart for Quarterly Sales
// Description: This C# snippet uses Aspose.Cells to build a new workbook, populate it with quarterly sales data for three products, add a ColumnStacked chart, assign product values (B2:D5) as series and quarters (A2:A5) as categories, set the chart title to "Cumulative Sales by Quarter", and save the file as StackedColumnChart.xlsx.
// Keywords: Aspose.Cells stacked column chart C# | ColumnStacked chart example | cumulative sales chart Aspose.Cells | add chart to Excel workbook .NET | save Excel with chart Aspose | quarterly sales visualization
// Common Searches: Aspose.Cells how to create stacked column chart | C# example for cumulative sales chart in Excel | Add ColumnStacked chart with series and categories using Aspose.Cells | Save workbook with stacked column chart Aspose .NET | Stacked column chart code sample Aspose.Cells
// Developer Intent: Generate a stacked column chart that displays cumulative product sales for each quarter using Aspose.Cells in C#.
// Use Cases: Produce a quarterly sales report with a ready‑to‑present stacked column chart for management. | Automate the creation of Excel templates that visualize product performance per quarter. | Export raw sales data to Excel while embedding a pre‑formatted stacked column chart for downstream analysis.
// AI Prompts: Write C# code with Aspose.Cells to create a stacked column chart from a data range and set a custom title. | Show how to add data labels, axis titles, and legend customization to a stacked column chart in Aspose.Cells. | Explain the steps to convert a stacked column chart to a 100 % stacked column chart using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# snippet uses Aspose.Cells to build a new workbook, populate it with quarterly sales data for three products, add a ColumnStacked chart, assign product values (B2:D5) as series and quarters (A2:A5) as categories, set the chart title to "Cumulative Sales by Quarter", and save the file as StackedColumnChart.xlsx.
class StackedColumnChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with quarterly sales data
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["B1"].PutValue("Product A");
        sheet.Cells["C1"].PutValue("Product B");
        sheet.Cells["D1"].PutValue("Product C");

        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["C2"].PutValue(150);
        sheet.Cells["D2"].PutValue(100);

        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B3"].PutValue(130);
        sheet.Cells["C3"].PutValue(160);
        sheet.Cells["D3"].PutValue(110);

        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B4"].PutValue(140);
        sheet.Cells["C4"].PutValue(170);
        sheet.Cells["D4"].PutValue(120);

        sheet.Cells["A5"].PutValue("Q4");
        sheet.Cells["B5"].PutValue(150);
        sheet.Cells["C5"].PutValue(180);
        sheet.Cells["D5"].PutValue(130);

        // Add a stacked column chart (ChartType.ColumnStacked) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the series and the category (quarters)
        chart.NSeries.Add("B2:D5", true);          // Series values (products)
        chart.NSeries.CategoryData = "A2:A5";      // Category labels (quarters)

        // Set a descriptive title for the chart
        chart.Title.Text = "Cumulative Sales by Quarter";

        // Save the workbook with the chart
        workbook.Save("StackedColumnChart.xlsx", SaveFormat.Xlsx);
    }
}
