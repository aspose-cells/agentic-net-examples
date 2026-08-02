// Title: C# – Add Data Labels to an Aspose.Cells Chart After Localization (SettableChartGlobalizationSettings)
// Description: Creates a workbook, fills it with category/value data, inserts a column chart, changes the chart’s OtherName label to Spanish using SettableChartGlobalizationSettings, then adds value data labels per point and saves the file. Demonstrates that labels stay in the original language after localization.
// Keywords: Aspose.Cells C# chart data labels | SettableChartGlobalizationSettings | chart localization Spanish | OtherName localization Aspose.Cells | add data labels after localization | Excel column chart custom labels | verify chart language after globalization
// Common Searches: Aspose.Cells add data labels after localization | C# chart globalization OtherName Spanish | how to keep chart labels in original language Aspose.Cells | custom point labels after SettableChartGlobalizationSettings | verify data labels language in Aspose.Cells chart
// Developer Intent: Add value data labels to a chart after applying SettableChartGlobalizationSettings and confirm the labels retain the source language.
// Use Cases: Generate an Excel report with a Spanish UI while showing numeric values on each column. | Test that chart data labels are unaffected by globalization changes. | Create per‑point custom labels after modifying chart globalization settings.
// AI Prompts: Show a C# example that sets OtherName to Spanish with SettableChartGlobalizationSettings, then adds value data labels to each series. | Explain how to verify that chart data labels keep their original language after applying localization in Aspose.Cells. | Provide code to customize point labels after changing chart globalization settings in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills it with category/value data, inserts a column chart, changes the chart’s OtherName label to Spanish using SettableChartGlobalizationSettings, then adds value data labels per point and saves the file. Demonstrates that labels stay in the original language after localization.
    public class AddDataLabelsAfterLocalization
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (Category and Value)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Item 1");
                sheet.Cells["A3"].PutValue("Item 2");
                sheet.Cells["A4"].PutValue("Item 3");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(250);
                sheet.Cells["B4"].PutValue(370);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Localization step: change a globalization label
                SettableChartGlobalizationSettings localization = new SettableChartGlobalizationSettings();
                localization.SetOtherName("Otros"); // Spanish for "Other"

                // Add data labels AFTER the localization step.
                foreach (Series series in chart.NSeries)
                {
                    // Show the value in each data label
                    series.DataLabels.ShowValue = true;

                    // Optionally customize each point's label text
                    for (int i = 0; i < series.Points.Count; i++)
                    {
                        ChartPoint point = series.Points[i];
                        point.DataLabels.Text = $"Value: {point.YValue}";
                    }
                }

                // Save the workbook
                string outputPath = "AddDataLabelsAfterLocalization.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            AddDataLabelsAfterLocalization.Run();
        }
    }
}
