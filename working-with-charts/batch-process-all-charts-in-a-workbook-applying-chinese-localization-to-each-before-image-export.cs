using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLocalization
{
    // Custom globalization settings that provide Chinese labels for chart elements
    public class ChineseChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Chinese unit names
            return type switch
            {
                DisplayUnitType.Hundreds => "百",
                DisplayUnitType.Thousands => "千",
                DisplayUnitType.TenThousands => "万",
                DisplayUnitType.Millions => "百万",
                DisplayUnitType.Billions => "十亿",
                _ => base.GetAxisUnitName(type),
            };
        }

        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        public override string GetAxisTitleName()
        {
            return "轴标题";
        }

        public override string GetLegendIncreaseName()
        {
            return "增加";
        }

        public override string GetLegendDecreaseName()
        {
            return "减少";
        }

        public override string GetLegendTotalName()
        {
            return "总计";
        }

        public override string GetOtherName()
        {
            return "其他";
        }

        public override string GetSeriesName()
        {
            return "系列";
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Apply Chinese globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChineseChartGlobalizationSettings()
            };

            // Ensure output directory exists
            string outputDir = "ChartImages";
            Directory.CreateDirectory(outputDir);

            // Iterate through all worksheets
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];
                ChartCollection charts = sheet.Charts;

                // Iterate through all charts in the worksheet
                for (int chartIndex = 0; chartIndex < charts.Count; chartIndex++)
                {
                    Chart chart = charts[chartIndex];

                    // Recalculate chart layout after applying globalization (optional but safe)
                    chart.Calculate();

                    // Build a unique file name for each chart image
                    string imageFile = Path.Combine(outputDir,
                        $"Sheet{wsIndex + 1}_Chart{chartIndex + 1}.png");

                    // Export chart to PNG image (rule: ToImage(string, ImageType))
                    chart.ToImage(imageFile, ImageType.Png);
                }
            }

            // Optionally save the workbook with applied settings (lifecycle rule: save)
            string outputWorkbook = "output_localized.xlsx";
            workbook.Save(outputWorkbook);
        }
    }
}