// Title: C# – Verify Default English Chart Localization with Aspose.Cells ChartGlobalizationSettings
// Description: Creates a workbook, adds a column chart, and uses a fresh ChartGlobalizationSettings instance (no custom culture) to fetch the built‑in English labels such as Series, Chart Title, Increase, Decrease, Total, Axis Title, Other and the "Thousands" axis unit. The values are printed to the console and optionally applied to the chart before saving the file.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | default chart localization | C# chart example | Aspose.Cells .NET | English chart labels | chart title default text | axis unit name thousands | no custom culture | Aspose.Cells demo
// Common Searches: Aspose.Cells get default chart text C# | ChartGlobalizationSettings default English strings | how to retrieve chart labels without localization in Aspose.Cells | default series name Aspose.Cells chart | axis unit name thousands Aspose.Cells
// Developer Intent: Confirm that Aspose.Cells returns the built‑in English strings for chart elements when no globalization settings are supplied.
// Use Cases: Display the default English series name, chart title, and legend entries in a console log for verification. | Assign the retrieved default labels to a chart's title and axis titles before exporting the workbook. | Implement an automated test that asserts each default label matches the expected English value (e.g., "Series", "Chart Title").
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells and prints the default localization strings using ChartGlobalizationSettings. | Explain how ChartGlobalizationSettings determines the fallback English text when no culture is set in Aspose.Cells. | Generate a C# unit test that verifies ChartGlobalizationSettings.GetChartTitleName() returns "Chart Title".

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalizationDemo
{
    // Creates a workbook, adds a column chart, and uses a fresh ChartGlobalizationSettings instance (no custom culture) to fetch the built‑in English labels such as Series, Chart Title, Increase, Decrease, Total, Axis Title, Other and the "Thousands" axis unit. The values are printed to the console and optionally applied to the chart before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Do NOT assign any custom globalization settings.
            // Use the default ChartGlobalizationSettings to retrieve default English texts.
            ChartGlobalizationSettings defaultSettings = new ChartGlobalizationSettings();

            // Retrieve default English strings
            string seriesName = defaultSettings.GetSeriesName();               // Expected: "Series"
            string chartTitleName = defaultSettings.GetChartTitleName();       // Expected: "Chart Title"
            string legendIncrease = defaultSettings.GetLegendIncreaseName();   // Expected: "Increase"
            string legendDecrease = defaultSettings.GetLegendDecreaseName();   // Expected: "Decrease"
            string legendTotal = defaultSettings.GetLegendTotalName();         // Expected: "Total"
            string axisTitle = defaultSettings.GetAxisTitleName();             // Expected: "Axis Title"
            string otherName = defaultSettings.GetOtherName();                 // Expected: "Other"
            string axisUnit = defaultSettings.GetAxisUnitName(DisplayUnitType.Thousands); // Expected: "Thousands"

            // Output the retrieved default texts to the console
            Console.WriteLine("Default Series Name: " + seriesName);
            Console.WriteLine("Default Chart Title Name: " + chartTitleName);
            Console.WriteLine("Default Legend Increase Name: " + legendIncrease);
            Console.WriteLine("Default Legend Decrease Name: " + legendDecrease);
            Console.WriteLine("Default Legend Total Name: " + legendTotal);
            Console.WriteLine("Default Axis Title Name: " + axisTitle);
            Console.WriteLine("Default Other Name: " + otherName);
            Console.WriteLine("Default Axis Unit Name (Thousands): " + axisUnit);

            // Optionally assign these defaults to the chart to demonstrate they appear as expected
            chart.Title.Text = chartTitleName;
            chart.ValueAxis.Title.Text = axisTitle;
            chart.CategoryAxis.Title.Text = axisTitle;

            // Save the workbook (using the standard save method)
            workbook.Save("DefaultChartLocalizationDemo.xlsx");
        }
    }
}
