// Title: SetChartDataRange with a Named Range (SalesData) for Column Chart in Aspose.Cells .NET
// Description: Shows how to build a workbook, define the named range "SalesData" over A1:B4, add a column chart, and link the chart to that range using Chart.SetChartDataRange in C#.
// Keywords: Aspose.Cells SetChartDataRange | named range chart .NET | C# chart data source named range | Aspose.Cells column chart example | create named range Aspose.Cells | bind chart to named range
// Common Searches: Aspose.Cells SetChartDataRange example | C# bind chart to named range | how to use named ranges in Aspose.Cells charts | set chart source by name Aspose.Cells | create column chart from named range C#
// Developer Intent: Link a column chart to the predefined "SalesData" range via SetChartDataRange.
// Use Cases: Generate a report where the chart updates automatically when the SalesData range changes. | Reuse the same named range across multiple charts for consistent visualizations. | Simplify maintenance of dashboards by centralizing data in a named range.
// AI Prompts: Provide C# code that creates a named range and attaches it to a chart using Aspose.Cells. | Explain how to switch between row‑wise and column‑wise data binding with SetChartDataRange. | Describe the steps to refresh a chart after modifying the data inside a named range.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, define the named range "SalesData" over A1:B4, add a column chart, and link the chart to that range using Chart.SetChartDataRange in C#.
    public class SetChartDataRangeWithNamedRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart (A1:B4)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Create a named range called "SalesData" that refers to the data area
                AsposeRange dataRange = sheet.Cells.CreateRange("A1:B4");
                dataRange.Name = "SalesData";

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Assign the named range as the chart source (by column)
                chart.SetChartDataRange("SalesData", true);

                // Save the workbook
                string outputPath = "ChartWithNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetChartDataRangeWithNamedRange.Run();
        }
    }
}
