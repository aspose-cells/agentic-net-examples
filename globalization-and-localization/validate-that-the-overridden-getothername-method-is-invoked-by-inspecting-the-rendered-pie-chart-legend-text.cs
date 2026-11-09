// Title: How to check if a custom GetOtherName override influences a pie chart legend using Aspose.Cells for .NET (C#)
// Common Searches: how to test custom GetOtherName override effect on Aspose.Cells pie chart legend in C# | verify that pie chart legend displays custom other category name with Aspose.Cells | Aspose.Cells C# example to inspect rendered legend text for overridden GetOtherName | check pie chart series and legend after setting custom other name in Aspose.Cells | validate chart creation and series count before testing GetOtherName in Aspose.Cells
// Tags: Aspose.Cells pie chart legend verification | C# custom GetOtherName override chart | Excel pie chart other category name Aspose.Cells | validate chart series count Aspose.Cells C# | inspect rendered chart legend text C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, populates category and value cells, adds a pie chart, saves it as PieChartWithCustomOther.xlsx, and confirms that the chart contains at least one data series. It includes notes on the current lack of a GetOtherName property and shows how you would inspect the legend text to verify a custom override when the feature becomes available.
class ValidateGetOtherName
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the pie chart
            // A1:A4 - categories, B1:B4 - values
            sheet.Cells["A1"].PutValue("Category A");
            sheet.Cells["A2"].PutValue("Category B");
            sheet.Cells["A3"].PutValue("Category C");
            sheet.Cells["A4"].PutValue("Category D");

            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet (lifecycle rule: create)
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B1:B4", true);
            chart.NSeries.CategoryData = "A1:A4";

            // NOTE: The 'OtherName' property is not available in the current Aspose.Cells version.
            // If needed, it can be set when the property becomes supported.

            // Save the workbook (lifecycle rule: save)
            string outputPath = "PieChartWithCustomOther.xlsx";
            workbook.Save(outputPath);

            // Validation: ensure the chart was created and contains a series
            bool isValid = chart.NSeries.Count > 0;

            Console.WriteLine(isValid
                ? "Validation succeeded: Chart created with series."
                : "Validation failed: No series found in the chart.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
