// Title: Set a column chart’s data source to a populated cell range using Chart.SetChartDataRange in Aspose.Cells for .NET
// AI Prompts: Create a new workbook, fill cells A1:B4 with categories and values, add a column chart, and bind it to that range with series plotted by column via Chart.SetChartDataRange. | Modify an existing Aspose.Cells chart to reference a different cell block as its data source and specify whether series are plotted by rows or columns. | Generate an Excel file where the chart automatically reflects changes in the underlying A1:B4 data range using SetChartDataRange.
// Common Searches: Aspose.Cells how to bind chart to specific cell range in C# | C# example using Chart.SetChartDataRange for column chart | set series orientation when assigning data range to Aspose.Cells chart | programmatically update chart data source range in Aspose.Cells workbook | using SetChartDataRange to link worksheet data to a chart in .NET
// Tags: chart set data range Aspose.Cells .NET | Chart.SetChartDataRange column chart | populate worksheet cells for chart source Aspose | bind worksheet data to chart Aspose.Cells | series orientation column chart Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, populates cells A1 through B4 with category labels and numeric values, adds a column chart, and binds the chart to that range using Chart.SetChartDataRange with the series‑plotted‑by‑column flag set to true, then saves the file as ChartDataRangeDemo.xlsx.
class SetChartDataRangeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells with sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["A2"].PutValue("Cat1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Cat2");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Cat3");
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart's data source range using SetChartDataRange
        // The second argument (true) indicates that data series are plotted by column
        chart.SetChartDataRange("A1:B4", true);

        // Save the workbook to a file
        workbook.Save("ChartDataRangeDemo.xlsx");
    }
}
