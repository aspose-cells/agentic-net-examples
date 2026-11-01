// Title: Create a custom ChartGlobalizationSettings class in C# to localize the “Other” slice label of an Aspose.Cells pie chart
// Common Searches: how to localize the 'Other' slice label in an Aspose.Cells pie chart using C# | override GetOtherName in ChartGlobalizationSettings for French label in Aspose.Cells | apply custom chart globalization settings to a pie chart in Aspose.Cells C# example | Aspose.Cells pie chart display 'Other' category with custom text | set custom 'Other' label for small categories in Aspose.Cells chart
// Tags: chart globalization settings customization Aspose.Cells | custom GetOtherName implementation for pie chart | localized other slice label in Excel chart | Aspose.Cells C# example for pie chart localization | Excel workbook pie chart other category localization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Custom globalization settings for charts
// The example defines a MyChartGlobalizationSettings class that overrides GetOtherName to return the French word "Autre". It creates a workbook with category data, adds many small categories to trigger an "Other" slice, inserts a pie chart, applies the custom globalization settings via reflection to WorkbookSettings, and saves the file as ChartWithLocalizedOther.xlsx.
public class MyChartGlobalizationSettings : ChartGlobalizationSettings
{
    // Override the default "Other" label used in pie charts
    public override string GetOtherName()
    {
        // Return a localized string, e.g., French for "Other"
        return "Autre";
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Prepare data for a pie chart
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(10);

            // Add many small categories to trigger the "Other" slice
            for (int i = 5; i <= 20; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Cat{i}");
                sheet.Cells[$"B{i}"].PutValue(1);
            }

            // Insert a pie chart (Charts.Add returns the chart index)
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 5, 20, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B20", true);
            chart.NSeries.CategoryData = "A2:A20";

            // Apply the custom globalization settings if the property exists
            var settings = workbook.Settings;
            var prop = typeof(WorkbookSettings).GetProperty("ChartGlobalizationSettings");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(settings, new MyChartGlobalizationSettings());
            }

            // Save the workbook
            string outputPath = "ChartWithLocalizedOther.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
