// Title: Duplicate a Chart and Set a New Category Axis Range with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, clones the first chart, copies its series, assigns a new CategoryData range (e.g., A2:A5), and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart duplication | C# copy Excel chart | set chart category axis range | duplicate chart programmatically | Aspose.Cells NSeries CategoryData | clone Excel chart .NET | modify chart axis range Aspose.Cells | Aspose.Cells chart example | Excel chart copy C# | Aspose.Cells chart API
// Common Searches: how to duplicate a chart with Aspose.Cells C# | Aspose.Cells set CategoryData for copied chart | copy Excel chart and change category axis range .NET | C# Aspose.Cells clone chart example | duplicate chart and assign new categories Aspose.Cells
// Developer Intent: Copy an existing chart in a workbook and point its category axis to a different cell range.
// Use Cases: Create a secondary chart that uses the same data series but different labels for comparative analysis. | Automate report generation where each period requires a chart with its own category range. | Provide a visual summary alongside the original chart without manually recreating it.
// AI Prompts: Generate C# code that duplicates an Excel chart with Aspose.Cells and sets CategoryData to a user‑defined range. | Explain how to preserve chart formatting while copying series and changing the category axis in Aspose.Cells. | Show how to loop through multiple charts in a worksheet, duplicate each, and assign distinct category ranges programmatically.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, clones the first chart, copies its series, assigns a new CategoryData range (e.g., A2:A5), and saves the file using Aspose.Cells for .NET.
class DuplicateChartExample
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string outputPath = "WorkbookWithDuplicatedChart.xlsx";

            // Verify source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet (adjust index/name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart to duplicate
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No chart found on the worksheet.");
                return;
            }

            // Get the original chart (first chart in the collection)
            Chart originalChart = sheet.Charts[0];

            // Determine position for the duplicated chart
            int upperLeftRow = originalChart.ChartObject.UpperLeftRow;
            int upperLeftColumn = originalChart.ChartObject.UpperLeftColumn;

            // Use a reasonable size for the new chart (you can adjust as needed)
            int lowerRightRow = upperLeftRow + 15;
            int lowerRightColumn = upperLeftColumn + 5;

            // Add a new chart with the same type and position as the original chart
            int newChartIndex = sheet.Charts.Add(
                originalChart.Type,
                upperLeftRow,
                upperLeftColumn,
                lowerRightRow,
                lowerRightColumn);

            Chart duplicatedChart = sheet.Charts[newChartIndex];

            // Copy each series from the original chart to the duplicated chart
            foreach (Series srcSeries in originalChart.NSeries)
            {
                // Add the series values to the new chart (isVertical = true assumes column‑wise data)
                duplicatedChart.NSeries.Add(srcSeries.Values, true);

                // Optionally copy the series name
                duplicatedChart.NSeries[duplicatedChart.NSeries.Count - 1].Name = srcSeries.Name;
            }

            // Assign a new category axis range to the duplicated chart
            // Example: use cells A2:A5 on the same worksheet as the new categories
            duplicatedChart.NSeries.CategoryData = "A2:A5";

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
