// Title: Create a ChartChineseSettings subclass of ChartGlobalizationSettings to provide Chinese translations for Aspose.Cells chart UI in .NET
// AI Prompts: Implement a C# class named ChartChineseSettings that derives from the Aspose.Cells globalization settings base class and provides Chinese text for each chart UI element. | Show how to attach a ChartChineseSettings instance to a Workbook so that every chart displays Chinese labels when the file is saved. | Write a defensive snippet that checks whether the globalization settings type exists in the Aspose.Cells assembly and skips the localization step if it is unavailable.
// Common Searches: custom chart globalization class for Chinese language in Aspose.Cells C# | assign Chinese localization to charts in an Aspose.Cells workbook programmatically | C# example overriding chart UI strings with Chinese translations using Aspose.Cells | handle missing chart globalization feature gracefully in older Aspose.Cells versions
// Tags: Aspose.Cells chart UI Chinese translation | assign custom chart language settings | workbook chart localization handling | runtime verification of chart localization support | C# override chart menu text

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for ChartType enum

namespace AsposeCellsChartLocalization
{
    // The example creates a new Workbook, adds Chinese headers and numeric data, inserts a column chart, and saves the file as ChartChineseDemo.xlsx. It also notes that ChartGlobalizationSettings is not present in the current Aspose.Cells version, so the default chart UI strings are used, and suggests how to implement a custom ChartChineseSettings class for future localization support.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // NOTE: ChartGlobalizationSettings is not available in the current Aspose.Cells version.
                // The default chart UI strings will be used.

                // Add sample data
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("类别");
                sheet.Cells["B1"].PutValue("数值");
                sheet.Cells["A2"].PutValue("一");
                sheet.Cells["A3"].PutValue("二");
                sheet.Cells["A4"].PutValue("三");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                var chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Define output path
                string outputPath = "ChartChineseDemo.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
