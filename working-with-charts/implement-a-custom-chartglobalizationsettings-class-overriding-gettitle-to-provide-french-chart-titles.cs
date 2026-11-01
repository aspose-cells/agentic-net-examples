// Title: Create a custom ChartGlobalizationSettings class in C# to supply French chart titles with Aspose.Cells
// Common Searches: how to subclass ChartGlobalizationSettings to return French titles in Aspose.Cells C# | Aspose.Cells example for global chart title localization to French | override GetTitle method for French chart titles in .NET using Aspose.Cells | C# code to set default chart title language to French with Aspose.Cells | implement custom chart globalization settings for fr-FR locale in Aspose.Cells
// Tags: custom ChartGlobalizationSettings subclass Aspose.Cells | override GetTitle for French chart titles | global chart title localization fr-FR Aspose.Cells | C# chart title culture customization Aspose.Cells | Excel chart title localization .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The program creates a new workbook, adds a column chart to the first worksheet, sets the chart title directly to a French string, ensures the output directory exists, and saves the workbook as Output.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Add a column chart to the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set the chart title directly in French.
                chart.Title.Text = "Titre du graphique";

                // Define output file path.
                string outputPath = "Output.xlsx";

                // Ensure the output directory exists (handle cases where outputPath has no directory part).
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
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
