// Title: C# – Add Data Labels to an Aspose.Cells Chart After Applying Localization
// Description: The sample creates a workbook, populates it with category and numeric data, configures chart globalization so the “Other” category appears in Japanese, inserts a column chart, enables value and category data labels, prefixes each point label with Japanese text, and saves the file as DataLabelsAfterLocalization.xlsx. It shows how to confirm that data‑label text remains in the original language after localization.
// Keywords: Aspose.Cells data labels | chart localization | SettableChartGlobalizationSettings | C# column chart | custom point label | Japanese Other label | verify chart labels | Aspose.Cells .NET example | Excel chart globalization | data label prefix
// Common Searches: Aspose.Cells add data labels after globalization | C# chart localization with custom Other label | keep data label language after chart globalization | Aspose.Cells verify localized chart labels | set Japanese Other label in Aspose.Cells chart
// Developer Intent: Create a column chart, apply a Japanese “Other” label via chart globalization, turn on value and category data labels, add a Japanese prefix to each point’s label, and save the workbook to verify the labels stay in the source language.
// Use Cases: Produce a sales‑report workbook where the “Other” category is shown in Japanese and each column displays both the numeric value and a Japanese‑prefixed label. | Build a localized dashboard that programmatically adds data labels after setting chart globalization, ensuring the labels retain their original language for end‑users. | Automate regression testing of chart localization by assigning custom label text post‑globalization and exporting the workbook for visual inspection.
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart, sets the 'Other' category name to Japanese using SettableChartGlobalizationSettings, enables data labels, and adds a Japanese prefix to each point's label. | Explain how Aspose.Cells preserves custom data‑label text after applying chart globalization settings and demonstrate verification by saving the workbook. | Provide a step‑by‑step guide for adding data labels to an Aspose.Cells chart after localization, covering ShowValue, ShowCategoryName, and custom point label text in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, populates it with category and numeric data, configures chart globalization so the “Other” category appears in Japanese, inserts a column chart, enables value and category data labels, prefixes each point label with Japanese text, and saves the file as DataLabelsAfterLocalization.xlsx. It shows how to confirm that data‑label text remains in the original language after localization.
    public class DataLabelsAfterLocalization
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate sample data for the chart
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(250);
                sheet.Cells["B4"].PutValue(180);

                // -------------------------------------------------
                // 2. Apply chart globalization settings (e.g., set "Other" label)
                // -------------------------------------------------
                SettableChartGlobalizationSettings globalizationSettings = new SettableChartGlobalizationSettings();
                // Set a custom name for the "Other" label in the original language (e.g., Japanese)
                globalizationSettings.SetOtherName("その他"); // "Other" in Japanese

                // -------------------------------------------------
                // 3. Add a column chart and bind the data
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // 4. Enable data labels for each series
                // -------------------------------------------------
                foreach (Series series in chart.NSeries)
                {
                    series.DataLabels.ShowValue = true;          // Show the numeric value
                    series.DataLabels.ShowCategoryName = true;   // Show the category name

                    // -------------------------------------------------
                    // 5. After localization, set custom text for each point
                    // -------------------------------------------------
                    foreach (ChartPoint point in series.Points)
                    {
                        // Prepend a label in the original language
                        point.DataLabels.Text = $"元の: {point.YValue}";
                    }
                }

                // -------------------------------------------------
                // 6. Save the workbook
                // -------------------------------------------------
                workbook.Save("DataLabelsAfterLocalization.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelsAfterLocalization.Run();
        }
    }
}
