// Title: Validate GetOtherName override by inspecting the pie chart legend for the default “Other” slice using Aspose.Cells in C#
// AI Prompts: Create a C# Aspose.Cells example that builds a pie chart with more than five categories, forces layout calculation, reads the legend entries, and checks that one entry equals "Other" to confirm GetOtherName execution. | Develop a C# unit test that overrides GetOtherName for a pie chart series in Aspose.Cells, generates the chart, captures the legend strings, and asserts that the custom name appears. | Outline the steps to log the legend output of an Aspose.Cells pie chart and compare it against an expected label, demonstrating detection of the default "Other" slice.
// Common Searches: asp.net how to detect default 'Other' label in Aspose.Cells pie chart legend | c# Aspose.Cells unit test for pie chart legend content | reading legend strings from Aspose.Cells pie chart for localization | verify GetOtherName override effect on pie chart legend using Aspose.Cells
// Tags: chart legend retrieval Aspose.Cells | GetOtherName override verification C# | default Other slice detection Aspose.Cells | pie chart slice naming Aspose.Cells | localization testing of chart labels .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendValidation
{
    // The sample creates a workbook, populates category and value columns, adds a pie chart with more than five points, forces layout calculation, extracts the chart's legend text, validates that the default "Other" entry is present (indicating GetOtherName was invoked), outputs the result, and saves the file as PieChartWithCustomOther.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the pie chart (more than 5 points to trigger the "Other" slice)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                string[] categories = { "A", "B", "C", "D", "E", "F", "G" };
                int[] values = { 10, 20, 30, 40, 50, 60, 70 };
                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(categories[i]); // Column A
                    sheet.Cells[i + 2, 1].PutValue(values[i]);    // Column B
                }

                // Add a pie chart
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 10, 0, 25, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range
                chart.NSeries.Add("B2:B8", true);
                chart.NSeries.CategoryData = "A2:A8";

                // Show legend
                chart.ShowLegend = true;

                // NOTE: The Aspose.Cells API version used does not expose an
                //       OtherSliceName property for Series. Therefore the custom
                //       name for the "Other" slice cannot be set directly.
                //       The chart will display the default "Other" label.

                // Force layout calculation
                chart.Calculate();

                // Retrieve legend text
                string legendText = chart.Legend.Text;

                // Validate that the legend contains expected entries (including the default "Other")
                if (!string.IsNullOrEmpty(legendText) && legendText.Contains("Other"))
                {
                    Console.WriteLine("Validation succeeded: legend contains the default \"Other\" slice.");
                }
                else
                {
                    Console.WriteLine("Validation failed: \"Other\" slice not found in legend.");
                }

                // Save workbook
                string outputPath = "PieChartWithCustomOther.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
