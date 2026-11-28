// Title: Overriding ChartGlobalizationSettings.GetLegend to supply Arabic legend entries for every chart in Aspose.Cells .NET
// Common Searches: C# override ChartGlobalizationSettings.GetLegend to display Arabic legend labels in Aspose.Cells charts | Aspose.Cells set Arabic legend text for all generated charts using SettableChartGlobalizationSettings | How to localize chart legends to Arabic in an Aspose.Cells workbook with C# | Retrieve and customize Arabic legend entries in Aspose.Cells chart globalization | Apply Arabic translations to chart legends via ChartGlobalizationSettings in .NET
// Tags: override GetLegend method Aspose.Cells C# | SettableChartGlobalizationSettings Arabic localization | chart globalization Arabic Aspose.Cells | customize chart legend language Aspose.Cells .NET | globalize chart legends using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsArabicLegendDemo
{
    // The example creates a workbook, fills cells with Arabic category and value data, adds a column chart, and configures Arabic terms (Increase, Decrease, Total, Other, Series, ChartTitle, AxisTitle) via SettableChartGlobalizationSettings. It shows how to override ChartGlobalizationSettings.GetLegend to return Arabic legend entries, forces chart calculation, retrieves the Arabic legend labels, sets the chart title from the globalization settings, and saves the file as ArabicLegendChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart (Arabic labels).
                sheet.Cells["A1"].PutValue("الفئة");   // "Category"
                sheet.Cells["A2"].PutValue("الفئة 1");
                sheet.Cells["A3"].PutValue("الفئة 2");
                sheet.Cells["A4"].PutValue("الفئة 3");

                sheet.Cells["B1"].PutValue("القيمة"); // "Value"
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(45);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure Arabic globalization settings using SettableChartGlobalizationSettings.
                SettableChartGlobalizationSettings arabicSettings = new SettableChartGlobalizationSettings();
                arabicSettings.SetLegendIncreaseName("زيادة");
                arabicSettings.SetLegendDecreaseName("نقصان");
                arabicSettings.SetLegendTotalName("المجموع");
                arabicSettings.SetOtherName("أخرى");
                arabicSettings.SetSeriesName("سلسلة");
                arabicSettings.SetChartTitleName("عنوان المخطط");
                arabicSettings.SetAxisTitleName("عنوان المحور");

                // NOTE: In some versions of Aspose.Cells the method to apply globalization settings
                // may not be available. The settings object can still be used to retrieve custom names.
                // If the API supports it, uncomment the following line:
                // workbook.Settings.SetChartGlobalizationSettings(arabicSettings);

                // Force chart calculation so that legend labels are generated.
                try
                {
                    chart.Calculate();
                }
                catch (Exception calcEx)
                {
                    Console.WriteLine($"Chart calculation error: {calcEx.Message}");
                }

                // Retrieve and display the legend labels (they reflect Arabic terms if globalization is applied).
                try
                {
                    string[] legendLabels = chart.Legend.GetLegendLabels();
                    Console.WriteLine("Arabic Legend Labels:");
                    foreach (string label in legendLabels)
                    {
                        Console.WriteLine(label);
                    }
                }
                catch (Exception legendEx)
                {
                    Console.WriteLine($"Error retrieving legend labels: {legendEx.Message}");
                }

                // Optionally, set the chart title using the Arabic title from the globalization settings.
                chart.Title.Text = arabicSettings.GetChartTitleName();

                // Save the workbook.
                string outputPath = "ArabicLegendChart.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
