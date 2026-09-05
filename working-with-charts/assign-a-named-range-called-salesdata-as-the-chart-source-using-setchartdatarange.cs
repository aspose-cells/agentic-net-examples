// Title: Assign a named range as the data source for a column chart using SetChartDataRange in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, defines a named range called SalesData covering A1:B4, adds a column chart, and binds the chart to the SalesData range with SetChartDataRange. | Demonstrate how to call Chart.SetChartDataRange in Aspose.Cells to link a chart to an existing named range within a .NET workbook.
// Common Searches: Aspose.Cells C# bind chart to named range using SetChartDataRange | example of setting chart source to a defined name in Aspose.Cells .NET | C# Aspose.Cells column chart from named range SalesData | using SetChartDataRange with workbook names in Aspose.Cells | assign named range as chart data source Aspose.Cells tutorial
// Tags: Aspose.Cells SetChartDataRange named range | C# column chart from named range | define named range workbook Aspose.Cells | chart source binding SetChartDataRange | Excel workbook chart data source C#

using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, defines a named range "SalesData" covering cells A1:B4, adds a column chart, sets the chart's data source to the named range using SetChartDataRange, and saves the file as ChartWithNamedRange.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Define a named range called "SalesData" that includes the data area
        int nameIdx = workbook.Worksheets.Names.Add("SalesData");
        workbook.Worksheets.Names[nameIdx].RefersTo = "=Sheet1!$A$1:$B$4";

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIdx];

        // Assign the named range as the chart source using SetChartDataRange
        chart.SetChartDataRange("SalesData", true);

        // Save the workbook
        workbook.Save("ChartWithNamedRange.xlsx");
    }
}
