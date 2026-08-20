// Title: Log chart globalization method calls with a custom SettableChartGlobalizationSettings in Aspose.Cells for .NET
// Description: This example creates a LoggingChartGlobalizationSettings class that inherits from SettableChartGlobalizationSettings and overrides each label‑related method to write a console message before calling the base implementation. The custom class is assigned to workbook.Settings.GlobalizationSettings.ChartSettings, a column chart is built, chart.Calculate() forces label generation, legend entries are accessed, and the workbook is saved, enabling developers to see exactly which culture‑specific methods run during chart rendering.
// Keywords: Aspose.Cells | .NET | chart globalization | SettableChartGlobalizationSettings | logging | debug localization | culture-specific labels | chart legend debugging | axis unit name tracing | chart rendering diagnostics
// Common Searches: Aspose.Cells log chart globalization methods | debug chart label localization .NET | override SettableChartGlobalizationSettings example | trace GetLegendIncreaseName call Aspose.Cells | how to capture chart globalization calls | chart localization debugging tutorial
// Developer Intent: The developer wants to capture and display each culture‑specific label method invoked while a chart is rendered, to troubleshoot and verify localization behavior.
// Use Cases: Detect missing or incorrect translations by logging calls such as GetLegendIncreaseName, GetSeriesName, and GetAxisUnitName. | Confirm that the appropriate axis unit name is selected for various DisplayUnitType values during chart calculation. | Validate custom titles, series names, and legend entries when applying localized resources to a chart. | Create a reusable debugging tool for chart localization across multiple workbooks.
// AI Prompts: Generate a LoggingChartGlobalizationSettings class that writes method calls to a file instead of the console. | Show how to integrate this logging class into an existing Aspose.Cells project that uses multiple chart types. | Write a unit test that asserts each overridden globalization method is invoked when chart.Calculate() runs. | Provide a PowerShell script to parse the console output and produce a summary report of called localization methods.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDebugging
{
    // Custom globalization settings that log each method call
    // This example creates a LoggingChartGlobalizationSettings class that inherits from SettableChartGlobalizationSettings and overrides each label‑related method to write a console message before calling the base implementation. The custom class is assigned to workbook.Settings.GlobalizationSettings.ChartSettings, a column chart is built, chart.Calculate() forces label generation, legend entries are accessed, and the workbook is saved, enabling developers to see exactly which culture‑specific methods run during chart rendering.
    public class LoggingChartGlobalizationSettings : SettableChartGlobalizationSettings
    {
        public override string GetLegendIncreaseName()
        {
            Console.WriteLine("GetLegendIncreaseName called");
            return base.GetLegendIncreaseName();
        }

        public override string GetLegendDecreaseName()
        {
            Console.WriteLine("GetLegendDecreaseName called");
            return base.GetLegendDecreaseName();
        }

        public override string GetLegendTotalName()
        {
            Console.WriteLine("GetLegendTotalName called");
            return base.GetLegendTotalName();
        }

        public override string GetOtherName()
        {
            Console.WriteLine("GetOtherName called");
            return base.GetOtherName();
        }

        public override string GetSeriesName()
        {
            Console.WriteLine("GetSeriesName called");
            return base.GetSeriesName();
        }

        public override string GetChartTitleName()
        {
            Console.WriteLine("GetChartTitleName called");
            return base.GetChartTitleName();
        }

        public override string GetAxisTitleName()
        {
            Console.WriteLine("GetAxisTitleName called");
            return base.GetAxisTitleName();
        }

        public override string GetAxisUnitName(DisplayUnitType type)
        {
            Console.WriteLine($"GetAxisUnitName called with type: {type}");
            return base.GetAxisUnitName(type);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Quarterly Sales";

            // Apply custom globalization settings that log method calls
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new LoggingChartGlobalizationSettings()
            };

            // Force chart calculation to generate labels and invoke globalization methods
            chart.Calculate();

            // Access legend entries to ensure labels are retrieved
            var legendLabels = chart.Legend.GetLegendLabels();

            // Output retrieved legend labels (optional, just to demonstrate usage)
            Console.WriteLine("Legend Labels:");
            foreach (string label in legendLabels)
            {
                Console.WriteLine(label);
            }

            // Save the workbook
            workbook.Save("LoggingChartGlobalizationSettingsDemo.xlsx");
        }
    }
}
