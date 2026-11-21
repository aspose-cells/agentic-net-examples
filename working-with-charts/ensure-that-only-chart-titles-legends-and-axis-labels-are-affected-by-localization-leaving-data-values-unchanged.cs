// Title: How to localize Excel chart titles, legends, and axis labels with Aspose.Cells for .NET while preserving data values
// Common Searches: C# Aspose.Cells translate chart titles without changing series data | example code to localize Excel chart legends using Aspose.Cells .NET | preserve chart series values while changing axis titles in Aspose.Cells | apply custom localization to chart axis labels in an existing workbook with Aspose.Cells | localize Excel chart titles and legends programmatically in .NET
// Tags: chart title localization Aspose.Cells | legend text localization Aspose.Cells | axis title localization Aspose.Cells | preserve chart series values Aspose.Cells | Excel chart localization C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing Excel workbook, iterates through each worksheet and its charts, and applies a placeholder localization routine to the chart title, legend text, and the primary category and value axis titles, while deliberately leaving the chart series data untouched, then saves the updated file.
class ChartLocalization
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
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
                    // Localize chart title
                    if (chart.Title != null && !string.IsNullOrEmpty(chart.Title.Text))
                    {
                        chart.Title.Text = Localize(chart.Title.Text);
                    }

                    // Localize legend text if a custom text is set
                    if (chart.Legend != null && !string.IsNullOrEmpty(chart.Legend.Text))
                    {
                        chart.Legend.Text = Localize(chart.Legend.Text);
                    }

                    // Localize primary category axis title
                    Axis categoryAxis = chart.CategoryAxis;
                    if (categoryAxis != null && categoryAxis.Title != null && !string.IsNullOrEmpty(categoryAxis.Title.Text))
                    {
                        categoryAxis.Title.Text = Localize(categoryAxis.Title.Text);
                    }

                    // Localize primary value axis title
                    Axis valueAxis = chart.ValueAxis;
                    if (valueAxis != null && valueAxis.Title != null && !string.IsNullOrEmpty(valueAxis.Title.Text))
                    {
                        valueAxis.Title.Text = Localize(valueAxis.Title.Text);
                    }

                    // NOTE: Secondary axes are not available in the current Aspose.Cells version.
                    // Chart series values are left unchanged.
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder localization method – replace with real translation logic.
    static string Localize(string original)
    {
        // Example: prepend a marker to indicate localization.
        return "[LOC]" + original;
    }
}
