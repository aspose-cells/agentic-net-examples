// Title: Remove the legend from a column chart in Aspose.Cells for .NET to expand the plotting area
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and disables its legend to increase the plot region. | Show how to assign false to chart.ShowLegend in Aspose.Cells to hide the legend. | Write a C# example that builds an Excel workbook, adds sample data, inserts a column chart, hides the legend, and saves the file.
// Common Searches: Aspose.Cells C# expand column chart area by removing legend | how to hide legend in an Excel column chart using Aspose.Cells .NET | increase chart plot region by disabling legend with Aspose.Cells API
// Tags: Aspose.Cells chart legend suppression C# | column chart plot area expansion Aspose.Cells | chart legend visibility Aspose.Cells .NET | Excel chart formatting without legend Aspose | C# Aspose.Cells chart customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RemoveLegendExample
{
    // Creates a workbook, adds sample data, inserts a column chart, disables the legend to maximize the plot area, and saves the workbook as an Excel file.
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
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to maximize the plotting area
            chart.ShowLegend = false;

            // Save the workbook
            workbook.Save("ChartWithoutLegend.xlsx");
        }
    }
}
