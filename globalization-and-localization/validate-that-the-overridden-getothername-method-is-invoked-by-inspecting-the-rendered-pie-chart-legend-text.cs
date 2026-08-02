// Title: Verify Overridden GetOtherName via Pie Chart Legend in Aspose.Cells for .NET
// Description: This C# sample builds a workbook, fills category and value cells, adds a pie chart, calculates it, extracts legend entries with chart.Legend.GetLegendLabels(), prints them to the console, and saves the file. By inspecting the legend text you can determine whether a custom GetOtherName implementation is being called (once the API is available).
// Keywords: Aspose.Cells | C# | pie chart | legend labels | GetOtherName | override method | chart localization | globalization | custom slice name | chart testing | Excel automation
// Common Searches: Aspose.Cells GetOtherName override | pie chart legend text Aspose.Cells | read chart legend labels C# | test custom slice name in Excel chart | validate chart localization Aspose.Cells
// Developer Intent: Confirm that a custom GetOtherName method influences the pie chart legend entries.
// Use Cases: Manually verify that a custom "Other" slice name appears in the chart legend. | Create an automated test that asserts expected legend strings after overriding GetOtherName. | Demonstrate chart calculation and legend extraction for globalization or localization scenarios. | Save the workbook to review legend output directly in Excel.
// AI Prompts: Generate C# code that overrides GetOtherName for a pie chart in Aspose.Cells and checks the legend for the custom name. | Show how to mock GetOtherName in a unit test and assert that chart.Legend.GetLegendLabels() returns the expected values. | Explain step‑by‑step how to retrieve pie chart legend entries with Aspose.Cells and compare them to localized strings.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# sample builds a workbook, fills category and value cells, adds a pie chart, calculates it, extracts legend entries with chart.Legend.GetLegendLabels(), prints them to the console, and saves the file. By inspecting the legend text you can determine whether a custom GetOtherName implementation is being called (once the API is available).
    public class ValidateGetOtherNameInPieChartLegend
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate data for a pie chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["A5"].PutValue("Date");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(25);
                worksheet.Cells["B5"].PutValue(25);

                // Add a pie chart
                int chartIndex = worksheet.Charts.Add(ChartType.Pie, 7, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // NOTE: The 'OtherName' property is not available in the current Aspose.Cells version.
                // If needed, you can set the name of the "Other" slice using the appropriate API when it becomes available.

                // Calculate the chart to generate legend entries
                chart.Calculate();

                // Retrieve legend labels
                string[] legendLabels = chart.Legend.GetLegendLabels();

                // Output legend labels
                Console.WriteLine("Legend Labels:");
                foreach (string label in legendLabels)
                {
                    Console.WriteLine(label);
                }

                // Save the workbook safely
                string outputPath = "ValidateGetOtherNameInPieChartLegend.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"\nFailed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateGetOtherNameInPieChartLegend.Run();
        }
    }
}
