using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    // Custom globalization settings for Chinese language
    // Override axis unit names to Chinese characters
    public class ChineseChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "百";
                case DisplayUnitType.Thousands:
                    return "千";
                case DisplayUnitType.TenThousands:
                    return "万";
                case DisplayUnitType.Millions:
                    return "百万";
                case DisplayUnitType.Billions:
                    return "十亿";
                default:
                    return base.GetAxisUnitName(type);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\InputXlsx";
            // Folder where the PNG images will be saved
            string outputFolder = @"C:\ChartPng";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all XLSX files in the source folder
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Load the workbook
                Workbook workbook = new Workbook(xlsxPath);

                // Apply Chinese globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ChineseChartGlobalizationSettings()
                };

                // Process each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Process each chart in the worksheet
                    for (int i = 0; i < sheet.Charts.Count; i++)
                    {
                        Chart chart = sheet.Charts[i];

                        // Build a unique file name: <WorkbookName>_<SheetName>_Chart<index>.png
                        string workbookName = Path.GetFileNameWithoutExtension(xlsxPath);
                        string sheetName = sheet.Name;
                        string chartFileName = $"{workbookName}_{sheetName}_Chart{i + 1}.png";

                        string chartFilePath = Path.Combine(outputFolder, chartFileName);

                        // Export the chart to PNG using the ToImage method (string, ImageType)
                        chart.ToImage(chartFilePath, ImageType.Png);
                    }
                }
            }

            Console.WriteLine("All charts have been exported to PNG with Chinese localization.");
        }
    }
}