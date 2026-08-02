// Title: C# – Extract the Localized “Other” Legend Label from a Pie Chart Using Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook with category data that groups small slices into an "Other" segment, renders a pie chart, calculates it, retrieves the legend entries, obtains the culture‑specific "Other" name via ChartGlobalizationSettings, and extracts that label for localization validation.
// Keywords: Aspose.Cells C# pie chart | extract Other label | chart legend localization | ChartGlobalizationSettings GetOtherName | .NET chart globalization | pie chart legend extraction | localization quality check Aspose.Cells
// Common Searches: Aspose.Cells get "Other" legend entry from pie chart | C# extract localized Other label Aspose.Cells | how to read pie chart legend after chart.Calculate | ChartGlobalizationSettings other name example | verify Other slice translation in Aspose.Cells
// Developer Intent: Retrieve the culture‑specific "Other" legend entry from a rendered pie chart for localization testing.
// Use Cases: Confirm that the automatically generated "Other" slice appears with the correct translation in different cultures. | Automate regression tests that compare extracted legend labels against expected localized strings. | Generate a report of all legend entries across multiple charts to detect missing or incorrect "Other" labels.
// AI Prompts: Write C# code with Aspose.Cells that extracts the "Other" legend label from a pie chart after calling chart.Calculate(). | Explain the role of ChartGlobalizationSettings.GetOtherName() and how to match it with legend entries. | Show how to loop through every chart in a workbook and collect each chart's localized "Other" label.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook with category data that groups small slices into an "Other" segment, renders a pie chart, calculates it, retrieves the legend entries, obtains the culture‑specific "Other" name via ChartGlobalizationSettings, and extracts that label for localization validation.
    public class ExtractOtherLabelFromPieChart
    {
        // Entry point required by the project
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data with many categories; small values will be grouped into "Other"
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");

                string[] categories = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
                double[] values = { 5000, 3000, 200, 150, 120, 100, 80, 70, 60, 50 }; // small values trigger "Other"

                for (int i = 0; i < categories.Length; i++)
                {
                    worksheet.Cells[i + 2, 0].PutValue(categories[i]);   // Column A
                    worksheet.Cells[i + 2, 1].PutValue(values[i]);      // Column B
                }

                // Add a pie chart
                int chartIndex = worksheet.Charts.Add(ChartType.Pie, 15, 0, 30, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set data range for the series and categories
                chart.NSeries.Add("B2:B11", true);
                chart.NSeries.CategoryData = "A2:A11";

                // Calculate the chart to render legend and other elements
                chart.Calculate();

                // Retrieve legend labels after calculation
                string[] legendLabels = chart.Legend.GetLegendLabels();

                // Get the localized "Other" name from globalization settings
                ChartGlobalizationSettings globalizationSettings = new ChartGlobalizationSettings();
                string otherName = globalizationSettings.GetOtherName();

                // Find the "Other" label in the legend (if present)
                string extractedOtherLabel = null;
                foreach (string label in legendLabels)
                {
                    if (label.Equals(otherName, StringComparison.OrdinalIgnoreCase))
                    {
                        extractedOtherLabel = label;
                        break;
                    }
                }

                // Output the result
                if (extractedOtherLabel != null)
                {
                    Console.WriteLine($"Extracted \"Other\" label: {extractedOtherLabel}");
                }
                else
                {
                    Console.WriteLine($"\"Other\" label not found in legend. Available labels:");
                    foreach (string label in legendLabels)
                    {
                        Console.WriteLine($"- {label}");
                    }
                }

                // Save the workbook (optional, for visual verification)
                workbook.Save("PieChartWithOtherLabel.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during execution: {ex.Message}");
            }
        }
    }
}
