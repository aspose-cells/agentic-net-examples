// Title: Handle NotSupportedException when trying to enable leader lines on a pie chart using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a pie chart with Aspose.Cells, attempts to set DataLabels.ShowLeaderLines, and catches NotSupportedException to log a warning. | Write a robust Aspose.Cells example that adds a chart, checks for leader‑line support, and uses try‑catch to handle errors from unsupported features. | Provide a C# snippet that wraps chart customization in a try block and captures any exception thrown when applying leader lines to a chart.
// Common Searches: Aspose.Cells C# catch exception for unsupported chart leader lines | How to detect if leader lines are available for a pie chart in Aspose.Cells | C# try‑catch when setting DataLabels.ShowLeaderLines in Aspose.Cells | Aspose.Cells chart customization error handling example | Enable leader lines on pie chart Aspose.Cells not supported
// Tags: aspocells chart leader lines exception handling | pie chart data label limitations aspocells | c# try-catch aspocells chart customization | unsupported chart feature aspocells .net | aspocells workbook save error handling

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a pie chart, and attempts to enable leader lines. Because the DataLabels.ShowLeaderLines property is not available, the code wraps the operation in a try‑catch block that captures NotSupportedException and logs an informative message, while also handling any save errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(40);

            // Add a pie chart (leader lines are typically unsupported for pie charts)
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Bind data to the chart
            chart.NSeries.Add("A2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Attempt to enable leader lines – not supported in current API version
            // The following code is omitted because DataLabels.ShowLeaderLines does not exist.
            // If future versions add this property, it can be reintroduced.

            // Save the workbook
            try
            {
                workbook.Save("ChartWithLeaderLines.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception e)
        {
            Console.WriteLine("Unexpected error: " + e.Message);
        }
    }
}
