// Title: Bind a Column Chart to the Range A1:B12 Using Chart.SetChartDataRange in Aspose.Cells for .NET
// AI Prompts: Generate a new workbook, fill cells A1:B12 with headers and numeric values, add a column chart, and call Chart.SetChartDataRange with the verticalSeries flag set to true, then save as XLSX. | Using Aspose.Cells in C#, programmatically create sample data in A1:B12, insert a column chart, bind the chart to that range via SetChartDataRange, and export the file. | Write C# code that populates a worksheet, creates a column chart, assigns its data source to A1:B12 by invoking SetChartDataRange(true), and writes the workbook to disk.
// Common Searches: Aspose.Cells C# set chart data source to A1:B12 | how to bind a column chart to a specific range with Chart.SetChartDataRange | vertical series flag true Chart.SetChartDataRange example Aspose.Cells | programmatically create chart and assign data range in .NET using Aspose.Cells | sample code for Chart.SetChartDataRange in C#
// Tags: Chart.SetChartDataRange binding column chart | populate worksheet range for chart data Aspose.Cells | vertical series flag true Aspose.Cells chart | create column chart programmatically C# | export workbook with chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells A1:B12 with sample data, adds a column chart, binds the chart to that range vertically using Chart.SetChartDataRange, and saves the workbook as ChartWithDataRange.xlsx.
class SetChartDataRangeExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Populate the range A1:B12 with sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 12; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the chart to the specified data range A1:B12 (vertical series)
        chart.SetChartDataRange("A1:B12", true);

        // Save the workbook
        workbook.Save("ChartWithDataRange.xlsx");
    }
}
