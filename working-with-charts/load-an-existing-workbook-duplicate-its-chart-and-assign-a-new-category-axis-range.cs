// Title: Duplicate an Excel chart and set a new category axis range using Aspose.Cells for .NET
// Description: Loads a workbook, verifies a chart exists on the first worksheet, clones the chart with the same type and position, copies all series, assigns a new CategoryData range (e.g., D2:D10), updates the chart title, and saves the workbook as a new file. Includes basic error handling for missing files or charts.
// Keywords: Aspose.Cells | duplicate chart | C# chart copy | category axis range | Excel chart manipulation | chart series copy | set CategoryData | Aspose.Cells .NET | clone Excel chart
// Common Searches: Aspose.Cells duplicate chart C# | How to copy a chart and change its category axis in Aspose.Cells | Set new CategoryData for a chart using Aspose.Cells | Clone Excel chart with new categories Aspose.Cells | C# Aspose.Cells chart series duplication
// Developer Intent: Duplicate an existing chart in a workbook and assign a different category axis range.
// Use Cases: Create a copy of a sales chart and point it to a separate set of category labels for side‑by‑side comparison. | Generate region‑specific charts by reusing the original series data while swapping the category axis to region‑specific values. | Automate report generation where the original chart remains unchanged and a modified version with new categories is added for drill‑down analysis.
// AI Prompts: Write C# code with Aspose.Cells that clones a chart, copies all its series, sets a new CategoryData range, updates the title, and saves the workbook. | Show how to check for the presence of charts before cloning and handle cases where the new category range is missing or invalid. | Explain error‑handling strategies for file‑not‑found and chart‑not‑found scenarios when duplicating a chart with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, verifies a chart exists on the first worksheet, clones the chart with the same type and position, copies all series, assigns a new CategoryData range (e.g., D2:D10), updates the chart title, and saves the workbook as a new file. Includes basic error handling for missing files or charts.
class DuplicateChartExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index/name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the original chart
            Chart originalChart = sheet.Charts[0];

            // Duplicate the chart by adding a new chart with the same type and position
            int newChartIndex = sheet.Charts.Add(
                originalChart.Type,
                originalChart.ChartObject.UpperLeftRow,
                originalChart.ChartObject.UpperLeftColumn,
                originalChart.ChartObject.LowerRightRow,
                originalChart.ChartObject.LowerRightColumn);

            Chart duplicatedChart = sheet.Charts[newChartIndex];

            // Copy all series from the original chart to the duplicated chart
            for (int i = 0; i < originalChart.NSeries.Count; i++)
            {
                Series srcSeries = originalChart.NSeries[i];
                // The Values property holds the data range for the series (e.g., "B2:B10")
                duplicatedChart.NSeries.Add(srcSeries.Values, true);
            }

            // Assign a new category axis range to the duplicated chart
            // Example: use cells D2:D10 as the new categories
            duplicatedChart.NSeries.CategoryData = "D2:D10";

            // Optionally, modify other properties of the duplicated chart
            duplicatedChart.Title.Text = "Duplicated Chart with New Category Axis";

            // Save the workbook with the duplicated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
