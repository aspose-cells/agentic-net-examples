// Title: Check for Primary and Secondary Category Axes in an Aspose.Cells Chart (C#)
// Description: Creates a workbook, populates category and value data, adds a column chart, and uses chart.HasAxis(AxisType.Category, true) and chart.HasAxis(AxisType.Category, false) to detect the presence of primary and secondary category axes. The results are printed to the console and the workbook is saved as ChartAxisCheck.xlsx.
// Keywords: Aspose.Cells chart.HasAxis | C# chart axis detection | primary category axis Aspose.Cells | secondary category axis Aspose.Cells | AxisType.Category example | Aspose.Cells chart API | check chart axes .NET
// Common Searches: Aspose.Cells how to check if a chart has a category axis | chart.HasAxis primary secondary C# | detect axis existence in Aspose.Cells chart | Aspose.Cells chart axis presence example
// Developer Intent: Identify whether a chart contains primary and/or secondary category axes using the HasAxis method.
// Use Cases: Validate axis existence before applying custom formatting or adding a new axis. | Conditionally add a secondary category axis only when it is missing. | Log axis configuration for dynamic report generation or auditing.
// AI Prompts: Generate C# code that adds a secondary category axis to an Aspose.Cells chart only if chart.HasAxis(AxisType.Category, false) returns false. | Explain the behavior of chart.HasAxis for different AxisType values and the primary/secondary flag in Aspose.Cells. | Create a reusable C# method that returns a dictionary indicating the presence of primary and secondary category and value axes for any given chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisCheck
{
    // Creates a workbook, populates category and value data, adds a column chart, and uses chart.HasAxis(AxisType.Category, true) and chart.HasAxis(AxisType.Category, false) to detect the presence of primary and secondary category axes. The results are printed to the console and the workbook is saved as ChartAxisCheck.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Determine if a primary category axis exists
            bool hasPrimaryCategoryAxis = chart.HasAxis(AxisType.Category, true);
            // Determine if a secondary category axis exists (if applicable)
            bool hasSecondaryCategoryAxis = chart.HasAxis(AxisType.Category, false);

            // Output the results
            Console.WriteLine("Primary Category Axis exists: " + hasPrimaryCategoryAxis);
            Console.WriteLine("Secondary Category Axis exists: " + hasSecondaryCategoryAxis);

            // Save the workbook
            workbook.Save("ChartAxisCheck.xlsx");
        }
    }
}
