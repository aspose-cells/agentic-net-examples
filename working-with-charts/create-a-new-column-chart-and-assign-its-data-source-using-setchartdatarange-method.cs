// Title: Bind a Column Chart to a Cell Range with SetChartDataRange in Aspose.Cells for .NET (C#)
// Description: Demonstrates creating a workbook, populating A1:B4, adding a Column chart, and linking it to the range using Chart.SetChartDataRange(true) before saving as XLSX.
// Keywords: Aspose.Cells SetChartDataRange | C# column chart example | Aspose.Cells chart data source | bind chart to range .NET | programmatic Excel chart Aspose
// Common Searches: Aspose.Cells SetChartDataRange column chart C# | how to bind chart to cell range Aspose.Cells | chart.SetChartDataRange true meaning | create column chart with Aspose.Cells .NET | sample code for chart data range Aspose
// Developer Intent: Generate a column chart and attach it to a predefined worksheet range using SetChartDataRange.
// Use Cases: Automated sales dashboards where monthly figures are visualized via a column chart linked to a dynamic range. | Template workbooks that include a pre‑styled chart automatically populated from user‑entered data. | Scheduled reporting scripts that refresh chart sources before exporting to XLSX for distribution.
// AI Prompts: Write C# code with Aspose.Cells to add a stacked column chart and set its data range to "A1:C10" plotted by rows. | Explain the effect of the boolean flag in Chart.SetChartDataRange on column‑wise versus row‑wise plotting. | Provide a guide to programmatically update an existing chart’s data source based on a variable range in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // Demonstrates creating a workbook, populating A1:B4, adding a Column chart, and linking it to the range using Chart.SetChartDataRange(true) before saving as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Cat1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Cat2");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cat3");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Assign the data source to the chart using SetChartDataRange
            // The second parameter (true) indicates that the data is plotted by column
            chart.SetChartDataRange("A1:B4", true);

            // Optional: set a title for clarity
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook to an XLSX file
            workbook.Save("ColumnChartWithDataRange.xlsx", SaveFormat.Xlsx);
        }
    }
}
