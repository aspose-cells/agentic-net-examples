// Title: Check for Primary and Secondary Category Axes in an Aspose.Cells Chart (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, sets its data range, and uses Chart.HasAxis(AxisType.Category, true) and Chart.HasAxis(AxisType.Category, false) to determine whether the chart has a primary and/or secondary category axis, then prints the results and saves the file.
// Keywords: Aspose.Cells | Chart.HasAxis | Category axis detection | primary axis | secondary axis | C# | .NET | Excel chart | axis existence check | Aspose.Cells chart example
// Common Searches: Aspose.Cells how to check if a chart has a category axis | C# Chart.HasAxis primary category axis | detect secondary category axis Aspose.Cells | verify chart axes before formatting Aspose | Aspose.Cells chart axis existence example
// Developer Intent: Identify whether a chart contains primary and secondary category axes using the HasAxis method.
// Use Cases: Validate chart structure before applying custom formatting or data labels. | Add a secondary category axis only when it does not already exist. | Adapt dynamic report generation based on axis presence. | Log axis configuration for auditing automated chart creation.
// AI Prompts: Generate C# code that adds a secondary category axis to an Aspose.Cells chart only if Chart.HasAxis(AxisType.Category, false) returns false. | Explain the parameters of Chart.HasAxis and how they differ for primary versus secondary axes in Aspose.Cells. | Write a reusable method that returns a dictionary with keys 'PrimaryCategory' and 'SecondaryCategory' indicating axis presence for any Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisCheck
{
    // Creates a workbook, adds sample data, inserts a column chart, sets its data range, and uses Chart.HasAxis(AxisType.Category, true) and Chart.HasAxis(AxisType.Category, false) to determine whether the chart has a primary and/or secondary category axis, then prints the results and saves the file.
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

            // Add a column chart to the worksheet
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
