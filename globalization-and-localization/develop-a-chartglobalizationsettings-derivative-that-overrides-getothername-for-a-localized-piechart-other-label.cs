// Title: C# – Localize the “Other” slice label in an Aspose.Cells pie chart with a custom ChartGlobalizationSettings
// Description: This example shows how to subclass ChartGlobalizationSettings, override GetOtherName to return a localized label (e.g., Spanish “Otros”), and apply the custom settings to a workbook that contains a pie chart. The code creates sample data, builds the chart, assigns the globalization settings via workbook.Settings.GlobalizationSettings.ChartSettings, and saves the Excel file.
// Keywords: Aspose.Cells ChartGlobalizationSettings | override GetOtherName | pie chart other label localization | C# Aspose.Cells example | Spanish chart label | custom chart globalization | Excel pie chart other slice | Aspose.Cells GitHub sample | globalization settings Aspose
// Common Searches: Aspose.Cells change Other label in pie chart | How to localize chart labels in Aspose.Cells .NET | Custom ChartGlobalizationSettings C# example | Set Spanish label for Other slice Aspose.Cells | Override GetOtherName for chart globalization
// Developer Intent: Create a subclass of ChartGlobalizationSettings that overrides GetOtherName to supply a localized label for the aggregated “Other” slice in a pie chart and apply it to a workbook.
// Use Cases: Generate Excel reports with pie charts that display the “Other” category in the target language (e.g., Spanish, French). | Apply the same custom globalization across multiple workbooks to keep chart terminology consistent. | Switch between different localization subclasses at runtime based on the user's locale. | Meet regional language compliance requirements in automated spreadsheet generation.
// AI Prompts: Write C# code that defines a CustomChartGlobalizationSettings class returning a French label for the "Other" slice and uses it in a bar chart. | Explain how to detect the user's culture and load the appropriate ChartGlobalizationSettings subclass in Aspose.Cells. | Provide steps to unit‑test that the overridden GetOtherName value appears correctly in the saved Excel file. | Show how to register multiple custom ChartGlobalizationSettings in a single workbook for different chart types.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Derive from ChartGlobalizationSettings and override GetOtherName
// This example shows how to subclass ChartGlobalizationSettings, override GetOtherName to return a localized label (e.g., Spanish “Otros”), and apply the custom settings to a workbook that contains a pie chart. The code creates sample data, builds the chart, assigns the globalization settings via workbook.Settings.GlobalizationSettings.ChartSettings, and saves the Excel file.
public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
{
    // Return a localized label for the "Other" slice in a pie chart
    public override string GetOtherName()
    {
        // Example: Spanish localization
        return "Otros";
    }
}

public class ChartOtherLabelDemo
{
    public static void Run()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a pie chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["A5"].PutValue("D");
        worksheet.Cells["A6"].PutValue("E");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(5);
        worksheet.Cells["B5"].PutValue(3);
        worksheet.Cells["B6"].PutValue(2);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 7, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B6", true);          // Values
        chart.NSeries.CategoryData = "A2:A6";      // Categories
        chart.Title.Text = "Sample Pie Chart";

        // Apply the custom globalization settings (lifecycle: load)
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new CustomChartGlobalizationSettings()
        };

        // Save the workbook (lifecycle: save)
        workbook.Save("PieChartOtherLabelDemo.xlsx");
    }
}

// Entry point for demonstration
class Program
{
    static void Main()
    {
        ChartOtherLabelDemo.Run();
    }
}
