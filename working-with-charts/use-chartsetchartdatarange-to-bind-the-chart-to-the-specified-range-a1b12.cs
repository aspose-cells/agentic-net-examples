// Title: Binding a column chart to the range A1:B12 with Chart.SetChartDataRange in Aspose.Cells (.NET)
// AI Prompts: Write C# code that creates a workbook, adds a column chart, and links the chart to the range A1:B12 using SetChartDataRange with vertical orientation. | Update an existing Aspose.Cells worksheet to reassign a chart's data source to A1:B12 by calling SetChartDataRange, then save the file.
// Common Searches: Aspose.Cells C# set chart data source to A1:B12 | How to use SetChartDataRange for a column chart in .NET | Bind chart to a vertical data range in Aspose.Cells example | C# Aspose.Cells tutorial for chart data range binding
// Tags: Aspose.Cells SetChartDataRange usage | column chart data source binding Aspose.Cells | vertical orientation chart data Aspose.Cells .NET | saving workbook after chart binding Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new workbook, optionally fills cells A1:B12, adds a column chart, binds the chart to that range with SetChartDataRange (vertical orientation), and saves the workbook as ChartDataRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // (Optional) Populate some sample data in A1:B12
            // for (int i = 0; i < 12; i++)
            // {
            //     sheet.Cells[i, 0].PutValue(i + 1);          // Column A
            //     sheet.Cells[i, 1].PutValue((i + 1) * 10); // Column B
            // }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart to the data range A1:B12 (vertical orientation)
            chart.SetChartDataRange("A1:B12", true);

            // Save the workbook to a file
            workbook.Save("ChartDataRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
