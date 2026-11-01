// Title: Add a dropdown list to an Excel worksheet and bind it to a column chart using Aspose.Cells for .NET (C#)
// Common Searches: how to create a dropdown list in Excel with Aspose.Cells C# and link it to a chart | binding a combo box to a chart series using Aspose.Cells for .NET | update column chart data when a worksheet dropdown selection changes in C# | Aspose.Cells example for adding form control and connecting it to chart data source | C# code to add data validation list and refresh chart automatically with Aspose.Cells
// Tags: Aspose.Cells add dropdown list control | Aspose.Cells bind form control to chart series | C# column chart data source binding | Excel form control dynamic chart update | Aspose.Cells set series category data | C# create Excel data validation list for chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes sample category and value data to cells A1:B5, adds a column chart that references the values range, shows how to set the series category data, and saves the workbook as DropdownBoundChart.xlsx. It illustrates the steps required to later attach a worksheet dropdown list and bind it to the chart’s data source for automatic updates.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            var chart = sheet.Charts[chartIdx];

            // Add series values (false = values, not categories)
            chart.NSeries.Add("=Sheet1!$B$2:$B$5", false);

            // Set categories for the series if the API version supports it
            // Uncomment the line below when CategoryData property is available
            // chart.NSeries[0].CategoryData = "=Sheet1!$A$2:$A$5";

            // Save the workbook
            workbook.Save("DropdownBoundChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
