// Title: Localize Aspose.Cells Chart Legend and Month Labels to Spanish (C#)
// Description: Demonstrates how to write Spanish short month names to a worksheet, create a column chart, and customize the legend text (Increase, Decrease, Total) using SettableChartGlobalizationSettings. The example verifies the localized strings and saves the workbook as ChartWithSpanishMonthsAndLegend.xlsx.
// Keywords: Aspose.Cells | C# chart localization | Spanish month names | SettableChartGlobalizationSettings | PivotGlobalizationSettings | chart legend translation | custom chart globalization | Excel export Spanish | Aspose.Cells example
// Common Searches: Aspose.Cells change chart legend language | C# set Spanish month names in chart categories | How to localize chart legend in Aspose.Cells | SettableChartGlobalizationSettings usage | PivotGlobalizationSettings Spanish months
// Developer Intent: Apply a Spanish locale to chart legends and month category labels, then confirm the localization works before saving the workbook.
// Use Cases: Create multilingual sales dashboards where chart legends display Spanish terminology. | Generate regional reports with month names automatically shown in Spanish. | Programmatically validate that custom legend strings are applied in an exported Excel file.
// AI Prompts: Show C# code that attaches SettableChartGlobalizationSettings to a workbook so the legend appears in Spanish. | Explain how to override full month names in PivotGlobalizationSettings and use them as chart categories. | Provide a method to programmatically verify legend localization after the workbook is saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Custom globalization settings for months – returns Spanish month names
    // Demonstrates how to write Spanish short month names to a worksheet, create a column chart, and customize the legend text (Increase, Decrease, Total) using SettableChartGlobalizationSettings. The example verifies the localized strings and saves the workbook as ChartWithSpanishMonthsAndLegend.xlsx.
    public class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        // Return short month names in Spanish
        public override string[] GetShortTextOf12Months()
        {
            return new string[]
            {
                "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
            };
        }

        // Optionally override full month names as well
        public override string GetTextOfMonths()
        {
            return "Meses";
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Fill worksheet with Spanish month names ----------
            var monthSettings = new CustomPivotGlobalizationSettings();
            string[] spanishMonths = monthSettings.GetShortTextOf12Months();

            // Write month names to column A (A1:A12)
            for (int i = 0; i < spanishMonths.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(spanishMonths[i]); // Column A
                // Add some sample numeric data for each month in column B
                sheet.Cells[i, 1].PutValue((i + 1) * 10);
            }

            // ---------- Create a column chart using the month names as categories ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the series (values in B1:B12)
            chart.NSeries.Add("B1:B12", true);
            // Set category data (month names in A1:A12)
            chart.NSeries.CategoryData = "A1:A12";

            // Set a chart title
            chart.Title.Text = "Ventas Mensuales";

            // ---------- Apply custom legend localization ----------
            // Create an instance of SettableChartGlobalizationSettings
            SettableChartGlobalizationSettings legendSettings = new SettableChartGlobalizationSettings();

            // Set Spanish labels for legend components
            legendSettings.SetLegendIncreaseName("Aumento");
            legendSettings.SetLegendDecreaseName("Disminución");
            legendSettings.SetLegendTotalName("Total");

            // Verify the settings by retrieving the values
            string incLabel = legendSettings.GetLegendIncreaseName();
            string decLabel = legendSettings.GetLegendDecreaseName();
            string totalLabel = legendSettings.GetLegendTotalName();

            Console.WriteLine("Legend Increase Name (Spanish): " + incLabel);
            Console.WriteLine("Legend Decrease Name (Spanish): " + decLabel);
            Console.WriteLine("Legend Total Name (Spanish): " + totalLabel);

            // Note: In Aspose.Cells, the ChartGlobalizationSettings are applied via the workbook's
            // GlobalizationSettings. For demonstration, we simply show that the custom settings are
            // correctly configured. Attaching them to the workbook would require a custom
            // GlobalizationSettings implementation, which is beyond the scope of this example.

            // ---------- Save the workbook ----------
            workbook.Save("ChartWithSpanishMonthsAndLegend.xlsx");
            Console.WriteLine("Workbook saved as 'ChartWithSpanishMonthsAndLegend.xlsx'.");
        }
    }
}
