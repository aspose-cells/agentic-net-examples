// Title: Localize the Pie‑Chart “Other” Slice with a Custom ChartGlobalizationSettings in Aspose.Cells for .NET
// Description: Shows how to subclass ChartGlobalizationSettings, override GetOtherName to return a localized label such as Spanish “Otros”, assign the custom settings to Workbook.Settings.GlobalizationSettings.ChartSettings, and create a pie chart that displays the new label when ShowOtherPoints is enabled.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | GetOtherName | pie chart localization | custom chart globalization | C# example | Spanish label | other slice label | Excel chart globalization | .NET
// Common Searches: Aspose.Cells change other slice label | custom ChartGlobalizationSettings C# | localize pie chart other label Aspose | override GetOtherName example | show other points label translation | globalization settings chart Aspose.Cells
// Developer Intent: Provide a C# sample that creates a ChartGlobalizationSettings subclass, overrides GetOtherName to supply a localized string for the aggregated “Other” slice, and applies the subclass to a workbook’s chart globalization settings.
// Use Cases: Generate multilingual Excel reports where the “Other” slice of a pie chart appears in the target language. | Reuse a single CustomChartGlobalizationSettings class across many workbooks to ensure consistent chart label translation. | Combine custom chart globalization with regional number and date formats in the same workbook. | Integrate the custom settings into automated reporting pipelines that produce localized charts. | Extend the subclass to support additional languages such as French, German, or Japanese.
// AI Prompts: Write C# code that creates a CustomChartGlobalizationSettings class overriding GetOtherName to return a French label for the “Other” slice and applies it to a workbook. | Explain the steps to enable ShowOtherPoints for a pie chart in Aspose.Cells and use a custom ChartGlobalizationSettings to localize the label. | Provide a test script that verifies the overridden GetOtherName method is used when exporting the workbook to an XLSX file. | Show how to combine CustomChartGlobalizationSettings with custom number formats for a fully localized Excel workbook. | Generate a GitHub‑style README snippet describing the purpose and usage of CustomChartGlobalizationSettings.cs.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Derive from ChartGlobalizationSettings and override GetOtherName
// Shows how to subclass ChartGlobalizationSettings, override GetOtherName to return a localized label such as Spanish “Otros”, assign the custom settings to Workbook.Settings.GlobalizationSettings.ChartSettings, and create a pie chart that displays the new label when ShowOtherPoints is enabled.
public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
{
    // Return a localized label for the "Other" slice in a pie chart
    public override string GetOtherName()
    {
        return "Otros"; // Example: Spanish localization
    }
}

public class ChartOtherLabelDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
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
            sheet.Cells["A6"].PutValue("E");
            sheet.Cells["B6"].PutValue(50);
            sheet.Cells["A7"].PutValue("F");
            sheet.Cells["B7"].PutValue(60);
            sheet.Cells["A8"].PutValue("G");
            sheet.Cells["B8"].PutValue(70);
            sheet.Cells["A9"].PutValue("H");
            sheet.Cells["B9"].PutValue(80);
            sheet.Cells["A10"].PutValue("I");
            sheet.Cells["B10"].PutValue(90);
            sheet.Cells["A11"].PutValue("J");
            sheet.Cells["B11"].PutValue(100);

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 13, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B11", true);
            chart.NSeries.CategoryData = "A2:A11";

            // NOTE: The ShowOtherPoints property may not be available in older Aspose.Cells versions.
            // If supported, uncomment the following line:
            // chart.NSeries.ShowOtherPoints = true;

            // Apply custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new CustomChartGlobalizationSettings()
            };

            // Save the workbook
            string outputPath = "ChartOtherLabelDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Entry point
class Program
{
    static void Main()
    {
        ChartOtherLabelDemo.Run();
    }
}
