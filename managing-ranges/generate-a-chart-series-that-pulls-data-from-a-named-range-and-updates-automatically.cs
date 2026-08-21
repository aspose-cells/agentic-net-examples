// Title: Create a dynamic column chart from named ranges in Aspose.Cells for .NET
// Description: Demonstrates how to define named ranges for values and categories, add a column chart, bind its series and category axis to those ranges, and save the workbook so the chart updates automatically when the source data changes.
// Keywords: Aspose.Cells C# chart named range | dynamic chart Aspose.Cells | auto‑updating column chart .NET | named range series Aspose.Cells | chart category axis named range | Aspose.Cells example C#
// Common Searches: Aspose.Cells bind chart series to named range | C# create chart that updates with named range | How to use named ranges in Aspose.Cells charts | Dynamic column chart Aspose.Cells .NET | Set chart category data from named range Aspose
// Developer Intent: Generate a column chart whose data series and category labels are linked to named ranges, ensuring the chart refreshes automatically when the underlying cells are modified.
// Use Cases: Sales dashboards that reflect real‑time figures without manual chart reconfiguration. | Template workbooks where chart ranges expand or contract based on data volume. | Automated reporting pipelines that produce up‑to‑date visualizations with a single code change.
// AI Prompts: Write C# code using Aspose.Cells to create a line chart that pulls values and categories from named ranges and updates automatically. | Explain how to change an existing chart's series to reference a different named range in Aspose.Cells for .NET. | Provide steps to rename a named range programmatically and keep the linked chart synchronized.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to define named ranges for values and categories, add a column chart, bind its series and category axis to those ranges, and save the workbook so the chart updates automatically when the source data changes.
class NamedRangeChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate sample data
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Sales");
        for (int i = 2; i <= 6; i++)
        {
            ws.Cells["A" + i].PutValue("Item " + (i - 1));
            ws.Cells["B" + i].PutValue(i * 10);
        }

        // Define a named range for the sales values
        int salesNameIdx = wb.Worksheets.Names.Add("SalesData");
        wb.Worksheets.Names[salesNameIdx].RefersTo = "=Sheet1!$B$2:$B$6";

        // Define a named range for the category labels
        int categoryNameIdx = wb.Worksheets.Names.Add("CategoryData");
        wb.Worksheets.Names[categoryNameIdx].RefersTo = "=Sheet1!$A$2:$A$6";

        // Add a column chart to the worksheet
        int chartIdx = ws.Charts.Add(ChartType.Column, 7, 0, 20, 8);
        Chart chart = ws.Charts[chartIdx];

        // Add a series that references the named range; it will update automatically when data changes
        chart.NSeries.Add("SalesData", true);

        // Set the category axis to use the named range for categories
        chart.NSeries.CategoryData = "CategoryData";

        // Optional: set a chart title
        chart.Title.Text = "Sales by Item";

        // Save the workbook
        wb.Save("NamedRangeChart.xlsx");
    }
}
