// Title: C# – Localize Excel chart titles, legends, and axis labels with Aspose.Cells while preserving data values
// AI Prompts: Write C# code that opens an existing .xlsx workbook using Aspose.Cells, iterates through each worksheet and chart, and replaces the text of the chart title, series (legend) names, and both axis titles via a custom Localize(string) method, then saves the modified file. | Create a C# example that shows how to apply a localization function only to chart textual elements (title, legend entries, axis titles) in an Excel file with Aspose.Cells, ensuring that all numeric data and series values stay unchanged.
// Common Searches: aspnet how to translate only chart titles and axis labels in Excel using Aspose.Cells | C# Aspose.Cells localize chart legend text without modifying data series | example code to iterate charts in a workbook and change chart text with Aspose.Cells | preserve chart data while updating chart captions in .xlsx via Aspose.Cells | apply custom localization function to Excel chart elements in C#
// Tags: Aspose.Cells chart text localization | C# Aspose.Cells modify chart titles | Aspose.Cells translate chart legends | Excel chart axis title localization C# | preserve chart data values Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLocalizationExample
{
    // The sample loads an existing workbook, walks through every worksheet and its charts, and uses a placeholder Localize method to replace the text of chart titles, series names (legends), and both category and value axis titles. Data values and series content remain untouched, and the workbook is saved to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // ----- Localize Chart Title -----
                        if (chart.Title != null && !string.IsNullOrEmpty(chart.Title.Text))
                        {
                            chart.Title.Text = Localize(chart.Title.Text);
                        }

                        // ----- Localize Legend (Series Names) -----
                        foreach (Series series in chart.NSeries)
                        {
                            if (!string.IsNullOrEmpty(series.Name))
                            {
                                series.Name = Localize(series.Name);
                            }
                        }

                        // ----- Localize Axis Titles -----
                        // Category (X) Axis Title
                        Axis categoryAxis = chart.CategoryAxis;
                        if (categoryAxis != null && categoryAxis.Title != null && !string.IsNullOrEmpty(categoryAxis.Title.Text))
                        {
                            categoryAxis.Title.Text = Localize(categoryAxis.Title.Text);
                        }

                        // Value (Y) Axis Title
                        Axis valueAxis = chart.ValueAxis;
                        if (valueAxis != null && valueAxis.Title != null && !string.IsNullOrEmpty(valueAxis.Title.Text))
                        {
                            valueAxis.Title.Text = Localize(valueAxis.Title.Text);
                        }

                        // ----- Localize Axis Tick Labels (if custom) -----
                        // Aspose.Cells does not expose a direct collection for custom tick labels.
                        // This placeholder demonstrates where such logic would be placed if needed.
                        if (categoryAxis != null && categoryAxis.TickLabels != null)
                        {
                            // Example: modify TickLabels.Font or other properties if required.
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to perform localization.
        // Replace this stub with actual localization logic (e.g., resource lookup, translation service).
        static string Localize(string text)
        {
            // Placeholder: return the original text.
            return text;
        }
    }
}
