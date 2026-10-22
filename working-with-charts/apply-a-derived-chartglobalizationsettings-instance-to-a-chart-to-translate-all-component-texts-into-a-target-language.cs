// Title: Derive and apply ChartGlobalizationSettings to translate all chart texts in an Aspose.Cells column chart (C#)
// AI Prompts: Write C# code that creates a column chart with sample data, derives a custom ChartGlobalizationSettings class to set a target language, assigns the instance to the chart's GlobalizationSettings, and saves the workbook as an XLSX file using Aspose.Cells. | Show how to subclass ChartGlobalizationSettings, override its language properties, and integrate the subclass with a chart object before exporting the workbook.
// Common Searches: how to use ChartGlobalizationSettings in Aspose.Cells to localize chart labels C# | example of custom ChartGlobalizationSettings subclass for translating Excel chart text | Aspose.Cells set chart language to French programmatically | apply globalisation settings to a column chart in .NET
// Tags: Aspose.Cells ChartGlobalizationSettings subclass example | C# chart text localization with Aspose.Cells | apply language translation to Excel chart components | column chart globalization Aspose.Cells .NET | export localized chart to XLSX using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills cells A1:B3 with month names and sales figures, adds a column chart, derives a ChartGlobalizationSettings subclass to specify a target language, applies this instance to the chart's GlobalizationSettings so that titles, axis labels, and legends are translated, and then saves the workbook as ChartGlobalization.xlsx while handling any exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Jan");
                sheet.Cells["B1"].PutValue(120);
                sheet.Cells["A2"].PutValue("Feb");
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["A3"].PutValue("Mar");
                sheet.Cells["B3"].PutValue(180);

                // Add a column chart (Add returns the chart index)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B1:B3", true);
                chart.NSeries.CategoryData = "A1:A3";
                chart.Title.Text = "Quarterly Sales";

                // Define output path
                string outputPath = "ChartGlobalization.xlsx";

                // Ensure the target directory exists (if a directory part is present)
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
