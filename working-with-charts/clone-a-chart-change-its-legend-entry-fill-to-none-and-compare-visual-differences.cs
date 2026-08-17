// Title: Clone an Aspose.Cells chart, remove legend fill, and compare PNG renders (C#)
// Description: Creates a workbook with sample data, adds a column chart, clones it, disables the legend entry fill on the clone using IsTextNoFill, renders both charts to PNG via MemoryStream, counts byte‑wise differences, and saves the workbook and images for visual inspection.
// Keywords: Aspose.Cells chart cloning C# | legend entry no fill Aspose.Cells | export chart to PNG memory stream | byte‑wise PNG comparison | Chart.Clone method | chart regression testing | C# Excel chart rendering | Aspose.Cells image options
// Common Searches: how to clone a chart in Aspose.Cells C# | remove legend fill from Aspose.Cells chart | compare two chart PNG files programmatically | export Aspose.Cells chart without writing to disk | Aspose.Cells Chart.Clone example
// Developer Intent: Copy an existing chart, make its legend transparent, render both original and cloned charts as PNG images, and quantify visual differences.
// Use Cases: Validate styling changes by generating before‑and‑after chart images. | Create a variant of a chart with a transparent legend for presentations while keeping the original unchanged. | Automate regression tests for chart rendering by comparing exported PNG byte arrays.
// AI Prompts: Generate C# code that clones an Aspose.Cells chart and sets the cloned chart's legend entry IsTextNoFill property to true. | Explain an efficient way to compare two PNG byte arrays produced by Aspose.Cells chart rendering and report the number of differing bytes. | Show alternative cloning techniques (e.g., Chart.Clone) in Aspose.Cells and discuss their impact on legend formatting.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook with sample data, adds a column chart, clones it, disables the legend entry fill on the clone using IsTextNoFill, renders both charts to PNG via MemoryStream, counts byte‑wise differences, and saves the workbook and images for visual inspection.
class ChartCloneAndCompare
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(200);

            // Add the original chart
            int originalChartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart originalChart = sheet.Charts[originalChartIdx];
            originalChart.NSeries.Add("B2:B3", true);
            originalChart.NSeries.CategoryData = "A2:A3";

            // Clone the original chart by creating a new chart with the same data
            int clonedChartIdx = sheet.Charts.Add(ChartType.Column, 20, 0, 30, 5);
            Chart clonedChart = sheet.Charts[clonedChartIdx];
            clonedChart.NSeries.Add("B2:B3", true);
            clonedChart.NSeries.CategoryData = "A2:A3";

            // Change the legend entry fill of the cloned chart to "no fill"
            if (clonedChart.NSeries.Count > 0)
            {
                LegendEntry clonedLegendEntry = clonedChart.NSeries[0].LegendEntry;
                clonedLegendEntry.IsTextNoFill = true;
            }

            // Export both charts to PNG images using MemoryStream
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions(); // default format is PNG

            using (MemoryStream originalImgStream = new MemoryStream())
            {
                originalChart.ToImage(originalImgStream, imgOptions);
                byte[] originalBytes = originalImgStream.ToArray();

                using (MemoryStream clonedImgStream = new MemoryStream())
                {
                    clonedChart.ToImage(clonedImgStream, imgOptions);
                    byte[] clonedBytes = clonedImgStream.ToArray();

                    // Simple byte‑by‑byte comparison to count differing bytes
                    int diffCount = 0;
                    int minLength = Math.Min(originalBytes.Length, clonedBytes.Length);
                    for (int i = 0; i < minLength; i++)
                    {
                        if (originalBytes[i] != clonedBytes[i])
                            diffCount++;
                    }
                    diffCount += Math.Abs(originalBytes.Length - clonedBytes.Length); // account for length difference

                    Console.WriteLine($"Number of differing bytes between original and cloned chart: {diffCount}");

                    // Save the workbook (contains both charts) and the two images for visual inspection
                    string workbookPath = "ChartCloneComparison.xlsx";
                    workbook.Save(workbookPath);

                    File.WriteAllBytes("OriginalChart.png", originalBytes);
                    File.WriteAllBytes("ClonedChart_NoFill.png", clonedBytes);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
