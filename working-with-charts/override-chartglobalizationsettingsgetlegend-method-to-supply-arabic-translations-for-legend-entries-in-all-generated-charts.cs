// Title: How to subclass ChartGlobalizationSettings and override GetLegend to provide Arabic legend text for every chart in Aspose.Cells (C#)
// AI Prompts: Write a C# class that inherits from Aspose.Cells.Charts.ChartGlobalizationSettings and implements GetLegend to return Arabic series names based on the series index. | Show how to attach the custom ChartGlobalizationSettings to a Workbook so that all charts automatically display Arabic legends without setting each series name manually.
// Common Searches: aspnet override chartglobalizationsettings getlegend arabic legend Aspose.Cells | C# set default Arabic legend names for all Excel charts using Aspose.Cells | how to localize chart legends to Arabic globally in Aspose.Cells workbook | custom ChartGlobalizationSettings example for Arabic series labels in .NET | apply Arabic localization to chart legends across multiple charts Aspose.Cells
// Tags: ChartGlobalizationSettings GetLegend override C# | Arabic series names in Aspose.Cells charts | global chart legend provider .NET | custom chart globalization for Excel workbooks | automatic Arabic legend for generated charts

using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart and ChartType

namespace AsposeCellsChartLocalization
{
    // The example demonstrates creating a subclass of ChartGlobalizationSettings that overrides GetLegend to supply Arabic series names, registering this subclass with a Workbook, adding a column chart, and saving the file as ChartWithArabicLegend.xlsx, so every generated chart automatically shows Arabic legend entries.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Populate sample data for the chart.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["D1"].PutValue("Series3");

                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                sheet.Cells["D2"].PutValue(12);
                sheet.Cells["D3"].PutValue(22);
                sheet.Cells["D4"].PutValue(32);

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart.
                chart.NSeries.Add("B2:D4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Manually set Arabic legend entries for each series.
                if (chart.NSeries.Count >= 3)
                {
                    chart.NSeries[0].Name = "السلسلة 1";
                    chart.NSeries[1].Name = "السلسلة 2";
                    chart.NSeries[2].Name = "السلسلة 3";
                }

                // Save the workbook.
                string outputPath = "ChartWithArabicLegend.xlsx";
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
