// Title: SetChartDataRange with a Named Range – Aspose.Cells C# Chart Example
// Description: Demonstrates how to create a workbook, define a named range called "SalesData" (A1:B5), add a column chart, and bind the chart to the named range using chart.SetChartDataRange("SalesData", true) before saving the file as ChartWithNamedRange.xlsx.
// Keywords: Aspose.Cells SetChartDataRange | named range chart source | C# chart from named range | Aspose.Cells column chart | Excel named range Aspose | chart data range example | SetChartDataRange C# | Aspose.Cells chart binding
// Common Searches: Aspose.Cells bind chart to named range C# | SetChartDataRange example .NET | create column chart using named range Aspose.Cells | how to use named ranges with charts in Aspose.Cells | chart data source named range Aspose.Cells
// Developer Intent: Use SetChartDataRange to assign the "SalesData" named range as the data source for a column chart in a .NET workbook.
// Use Cases: Generate a quarterly‑sales column chart by referencing a predefined named range that includes headers. | Reuse a single named range across multiple charts to keep data references consistent. | Update chart values automatically by editing cells inside the named range without changing chart code.
// AI Prompts: Show C# code that changes an existing Aspose.Cells chart to use a different named range. | Provide an example of a line chart that references a named range located on another worksheet using SetChartDataRange. | Explain how to read the current data range of a chart and replace it with a new named range in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a named range called "SalesData" (A1:B5), add a column chart, and bind the chart to the named range using chart.SetChartDataRange("SalesData", true) before saving the file as ChartWithNamedRange.xlsx.
    public class SetChartDataRangeWithNamedRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["A5"].PutValue("Q4");
                sheet.Cells["B5"].PutValue(200);

                // Create a named range called "SalesData" that refers to the data area (including headers)
                int nameIndex = workbook.Worksheets.Names.Add("SalesData");
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$B$5";

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart's data source to the named range "SalesData"
                chart.SetChartDataRange("SalesData", true);

                // Set chart title
                chart.Title.Text = "Quarterly Sales";

                // Save the workbook
                workbook.Save("ChartWithNamedRange.xlsx");
                Console.WriteLine("Workbook saved as ChartWithNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetChartDataRangeWithNamedRange.Run();
        }
    }
}
