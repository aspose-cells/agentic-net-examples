// Title: Clone and Rename an Excel Chart on Another Worksheet with Aspose.Cells for .NET (C#)
// Description: A complete C# example that creates a workbook, adds sample data and a column chart, then clones the chart onto a new worksheet, copies its series, updates the chart title, and saves the file. Demonstrates chart cloning, title modification, and worksheet handling using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart cloning | duplicate Excel chart | copy chart to another worksheet | change chart title Aspose.Cells | .NET Excel automation | chart series copy | Excel workbook example | GitHub Aspose.Cells sample | Aspose.Cells API
// Common Searches: Aspose.Cells clone chart C# | copy Excel chart to another sheet using Aspose.Cells | change chart title programmatically Aspose.Cells .NET | duplicate chart worksheet Aspose.Cells example | how to copy chart series with Aspose.Cells
// Developer Intent: The developer needs to programmatically duplicate an existing chart, modify its title, and place the copy on a different worksheet using Aspose.Cells for .NET.
// Use Cases: Create a summary sheet that aggregates charts from multiple source sheets with uniform titles. | Generate a master report where each department’s chart is cloned and renamed for consistent branding. | Automate side‑by‑side chart comparisons by copying a chart to a separate worksheet and assigning a new title.
// AI Prompts: Show me C# code to clone an Excel chart to another worksheet and set a new title with Aspose.Cells. | How can I copy all series from an existing chart to a new chart on a different sheet using Aspose.Cells for .NET? | Explain how to preserve chart formatting while duplicating a chart to another worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    // A complete C# example that creates a workbook, adds sample data and a column chart, then clones the chart onto a new worksheet, copies its series, updates the chart title, and saves the file. Demonstrates chart cloning, title modification, and worksheet handling using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Source worksheet with original chart ----------
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "SourceSheet";

                // Populate sample data
                srcSheet.Cells["A1"].PutValue("Category");
                srcSheet.Cells["B1"].PutValue("Value");
                srcSheet.Cells["A2"].PutValue("A");
                srcSheet.Cells["A3"].PutValue("B");
                srcSheet.Cells["A4"].PutValue("C");
                srcSheet.Cells["B2"].PutValue(10);
                srcSheet.Cells["B3"].PutValue(20);
                srcSheet.Cells["B4"].PutValue(30);

                // Add an original chart
                int srcChartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart srcChart = srcSheet.Charts[srcChartIdx];
                srcChart.NSeries.Add("B2:B4", true);
                // Category data can be omitted; default numeric categories will be used
                srcChart.Title.Text = "Original Chart";

                // ---------- Destination worksheet ----------
                Worksheet destSheet = workbook.Worksheets.Add("ClonedChartSheet");

                // Clone the chart: create a new chart with same type and position
                int clonedChartIdx = destSheet.Charts.Add(
                    srcChart.Type,
                    srcChart.ChartObject.UpperLeftRow,
                    srcChart.ChartObject.UpperLeftColumn,
                    srcChart.ChartObject.LowerRightRow,
                    srcChart.ChartObject.LowerRightColumn);
                Chart clonedChart = destSheet.Charts[clonedChartIdx];

                // Copy series from the source chart to the cloned chart
                foreach (Series series in srcChart.NSeries)
                {
                    // Add series values; true indicates that the series is a column series
                    clonedChart.NSeries.Add(series.Values, true);
                    // Category data is not set explicitly; default categories will be used
                }

                // Change the title of the cloned chart
                clonedChart.Title.Text = "Cloned Chart Title";

                // Save the workbook
                string outputPath = "ChartCloneDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{System.IO.Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
