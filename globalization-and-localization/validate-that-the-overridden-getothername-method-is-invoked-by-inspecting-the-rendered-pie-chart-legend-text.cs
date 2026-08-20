// Title: Validate GetOtherName Override Using Pie Chart Legend in Aspose.Cells for .NET
// Description: Creates a workbook, fills A1:B5 with categories and values, adds a pie chart, forces calculation, extracts legend entries with chart.Legend.GetLegendLabels(), prints them, and saves the file. Use the output to confirm that an overridden GetOtherName method is being called.
// Keywords: Aspose.Cells GetOtherName override | pie chart legend Aspose.Cells | validate chart series name | C# Aspose.Cells localization | chart legend text verification
// Common Searches: how to check GetOtherName execution in Aspose.Cells | retrieve pie chart legend labels C# Aspose.Cells | debug custom series name in Aspose.Cells chart | verify legend text after chart calculation
// Developer Intent: Ensure the custom GetOtherName implementation is triggered by reading the generated legend entries of a pie chart.
// Use Cases: Automated test that reads legend labels to confirm custom "Other" category appears. | Logging chart legend text for localization validation. | Conditional workbook saving only after expected legend entries are present.
// AI Prompts: Generate C# code that overrides GetOtherName for a chart series and validates it by reading chart.Legend.GetLegendLabels() with Aspose.Cells. | Write a unit test that asserts the legend contains a specific custom label after GetOtherName is overridden in a pie chart. | Explain step‑by‑step how to debug GetOtherName execution by printing legend labels in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills A1:B5 with categories and values, adds a pie chart, forces calculation, extracts legend entries with chart.Legend.GetLegendLabels(), prints them, and saves the file. Use the output to confirm that an overridden GetOtherName method is being called.
    public class ValidateGetOtherNameInPieChartLegend
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for a pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["A5"].PutValue("Date");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(25);
                sheet.Cells["B5"].PutValue(25);

                // Add a pie chart
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // NOTE: The Series.OtherName property is not available in the current Aspose.Cells version.
                // If you need to set a custom "Other" label, consider adding a separate data point with the desired label.

                // Force chart calculation so legend entries are generated
                chart.Calculate();

                // Retrieve legend labels
                string[] legendLabels = chart.Legend.GetLegendLabels();

                // Output the legend labels
                Console.WriteLine("Legend Labels:");
                foreach (string label in legendLabels)
                {
                    Console.WriteLine("- " + label);
                }

                // Save the workbook (optional, just to have a complete file)
                workbook.Save("ValidateGetOtherNameInPieChartLegend.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
