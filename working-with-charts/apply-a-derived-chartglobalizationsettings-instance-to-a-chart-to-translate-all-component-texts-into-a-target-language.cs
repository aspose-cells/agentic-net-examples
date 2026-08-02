// Title: Aspose.Cells C# – Localize All Chart Texts to Spanish with a Custom ChartGlobalizationSettings
// Description: Demonstrates how to subclass ChartGlobalizationSettings, override its methods to return Spanish strings for titles, legends, axis labels and display‑unit names, assign the instance to a workbook's GlobalizationSettings, create a column chart, enable axis unit translation, and save the file as an Excel workbook with every chart component rendered in Spanish.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | C# chart localization | Spanish Excel chart | globalization settings Aspose | translate chart labels | display unit translation | multi‑language reporting | Excel chart example | Aspose.Cells tutorial
// Common Searches: how to localize chart titles in Aspose.Cells .NET | custom ChartGlobalizationSettings example for Spanish | apply chart globalization settings to a workbook | translate chart legends and axis labels with Aspose.Cells | display unit language for Excel charts C#
// Developer Intent: Apply a derived ChartGlobalizationSettings object so that every textual element of an Aspose.Cells chart is automatically shown in a target language.
// Use Cases: Create sales dashboards for Spanish‑speaking users with fully localized chart titles, legends, and axis units. | Build a multi‑locale reporting engine that swaps different ChartGlobalizationSettings subclasses (e.g., French, German, Portuguese) before exporting workbooks. | Show localized display‑unit labels such as "Miles" for thousands when using chart axis display units.
// AI Prompts: Generate a C# example that defines a ChartGlobalizationSettings subclass for French and applies it to an Aspose.Cells chart. | Explain how to override GetAxisUnitName to provide custom translations for display units in Aspose.Cells charts. | Provide code to switch between several ChartGlobalizationSettings implementations at runtime based on the user's language preference.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartGlobalizationDemo
{
    // Custom globalization settings for charts – translates component texts to Spanish
    // Demonstrates how to subclass ChartGlobalizationSettings, override its methods to return Spanish strings for titles, legends, axis labels and display‑unit names, assign the instance to a workbook's GlobalizationSettings, create a column chart, enable axis unit translation, and save the file as an Excel workbook with every chart component rendered in Spanish.
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Chart title translation
        public override string GetChartTitleName()
        {
            return "Título del Gráfico";
        }

        // Series name translation
        public override string GetSeriesName()
        {
            return "Serie";
        }

        // Legend increase label translation
        public override string GetLegendIncreaseName()
        {
            return "Aumento";
        }

        // Legend decrease label translation
        public override string GetLegendDecreaseName()
        {
            return "Disminución";
        }

        // Legend total label translation
        public override string GetLegendTotalName()
        {
            return "Total";
        }

        // Axis title translation
        public override string GetAxisTitleName()
        {
            return "Título del Eje";
        }

        // "Other" label translation
        public override string GetOtherName()
        {
            return "Otro";
        }

        // Axis unit translation (example for thousands)
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "Cientos";
                case DisplayUnitType.Thousands:
                    return "Miles";
                case DisplayUnitType.TenThousands:
                    return "Decenas de Miles";
                default:
                    return base.GetAxisUnitName(type);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Categoría");
            sheet.Cells["A2"].PutValue("Enero");
            sheet.Cells["A3"].PutValue("Febrero");
            sheet.Cells["A4"].PutValue("Marzo");
            sheet.Cells["B1"].PutValue("Ventas");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(1800);

            // Apply custom chart globalization settings to the workbook
            GlobalizationSettings globalization = new GlobalizationSettings
            {
                ChartSettings = new CustomChartGlobalizationSettings()
            };
            workbook.Settings.GlobalizationSettings = globalization;

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Configure the chart data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // The chart title will be automatically localized using GetChartTitleName()
            chart.Title.Text = "Ventas Mensuales";

            // Enable display of axis unit label to see GetAxisUnitName in action
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Save the workbook
            workbook.Save("ChartWithSpanishGlobalization.xlsx");
        }
    }
}
