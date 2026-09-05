// Title: Set category axis tick label orientation to stacked (vertical) in an Aspose.Cells column chart using C#
// AI Prompts: Write C# code using Aspose.Cells that builds a column chart and applies a stacked text direction to the category axis tick labels. | Provide a snippet that sets ChartTextDirectionType.Stacked on the CategoryAxis.TickLabels property of a chart in Aspose.Cells for .NET. | Explain how to orient axis labels vertically and export the workbook to an .xlsx file with Aspose.Cells.
// Common Searches: rotate chart axis labels to stacked orientation using Aspose.Cells C# | set category axis labels to vertical layout in Aspose.Cells column chart C# | Aspose.Cells tutorial for setting tick label direction on chart axes | C# code sample for changing chart label orientation in Aspose.Cells
// Tags: Aspose.Cells chart axis label direction | stacked text direction for chart axis | C# column chart vertical label orientation | modify chart axis tick labels in Aspose.Cells | export workbook to xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, sets the category axis tick labels to a stacked (vertical) orientation using ChartTextDirectionType.Stacked, and saves the file as TickLabelsStackedDirection.xlsx.
class SetTickLabelsDirection
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set tick labels direction to Stacked (vertical orientation)
        chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Stacked;

        // Save the workbook
        workbook.Save("TickLabelsStackedDirection.xlsx");
    }
}
