// Title: Batch export Excel charts to PNG with Chinese localization using Aspose.Cells for .NET
// Description: Scans a folder of .xlsx files, loads each workbook with Aspose.Cells, applies a custom ChartGlobalizationSettings that supplies Chinese axis unit names, iterates through every worksheet and chart, and saves each chart as a uniquely named PNG in a target directory.
// Keywords: Aspose.Cells | C# chart export | Excel to PNG | batch chart conversion | Chinese chart globalization | ChartGlobalizationSettings | localize Excel charts | automated chart image generation | .NET Excel chart export | Chinese axis units
// Common Searches: export all charts from multiple Excel files to PNG using Aspose.Cells | apply Chinese axis unit names when exporting Excel charts .NET | batch process XLSX workbooks to generate localized chart images | C# code to convert Excel charts to PNG with Chinese localization | Aspose.Cells chart globalization settings for China
// Developer Intent: Automatically export every chart in each XLSX workbook within a folder to PNG files while applying Chinese localization for axis labels.
// Use Cases: Generate chart assets for a Chinese‑language reporting portal where axis labels must display Chinese numeric units. | Create localized chart images from a library of Excel templates for documentation, presentations, or marketing dashboards. | Automate the production of PNG graphics for a data‑driven website that serves users in Mainland China.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through all .xlsx files in a directory, apply a custom ChineseChartGlobalizationSettings to each workbook, and export every chart as a PNG file. | Explain how to customize axis unit names for Chinese in Aspose.Cells chart globalization and ensure the settings are applied during batch export. | Suggest improvements to handle sub‑folder recursion, add logging for each exported chart, and implement robust exception handling.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartExportExample
{
    // Custom globalization settings for Chinese language
    // Scans a folder of .xlsx files, loads each workbook with Aspose.Cells, applies a custom ChartGlobalizationSettings that supplies Chinese axis unit names, iterates through every worksheet and chart, and saves each chart as a uniquely named PNG in a target directory.
    public class ChineseChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Provide Chinese unit names for axis display units
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
        static void Main(string[] args)
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\InputXlsx";
            // Folder where the exported PNG images will be saved
            string outputFolder = @"C:\ExportedCharts";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the source folder (non‑recursive)
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Load the workbook
                Workbook workbook = new Workbook(xlsxPath);

                // Apply Chinese chart globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ChineseChartGlobalizationSettings()
                };

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    for (int i = 0; i < sheet.Charts.Count; i++)
                    {
                        Chart chart = sheet.Charts[i];

                        // Build a unique file name for each chart image
                        string chartFileName = $"{Path.GetFileNameWithoutExtension(xlsxPath)}_Sheet{sheet.Index}_Chart{i + 1}.png";
                        string chartFilePath = Path.Combine(outputFolder, chartFileName);

                        // Export the chart to PNG (extension determines format)
                        chart.ToImage(chartFilePath);
                    }
                }

                // Optionally, save the workbook back if any modifications are needed
                // workbook.Save(xlsxPath, SaveFormat.Xlsx);
            }

            Console.WriteLine("Chart export completed.");
        }
    }
}
