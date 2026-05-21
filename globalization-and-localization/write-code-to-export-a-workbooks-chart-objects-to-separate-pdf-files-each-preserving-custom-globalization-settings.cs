using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ExportChartsToPdf
{
    // Custom globalization settings for charts
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Example: customize axis unit names
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "Hundreds_Custom";
                case DisplayUnitType.Thousands:
                    return "Thousands_Custom";
                case DisplayUnitType.TenThousands:
                    return "TenThousands_Custom";
                default:
                    return base.GetAxisUnitName(type);
            }
        }

        // Additional overrides can be added as needed
    }

    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Apply custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new CustomChartGlobalizationSettings()
            };

            // Iterate through all worksheets
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];

                // Iterate through all charts in the current worksheet
                for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
                {
                    Chart chart = sheet.Charts[chartIndex];

                    // Build a unique PDF file name for each chart
                    string pdfFileName = $"Chart_Sheet{wsIndex + 1}_Chart{chartIndex + 1}.pdf";

                    // Export the chart to PDF while preserving the custom globalization settings
                    chart.ToPdf(pdfFileName);
                }
            }

            Console.WriteLine("All charts have been exported to separate PDF files.");
        }
    }
}