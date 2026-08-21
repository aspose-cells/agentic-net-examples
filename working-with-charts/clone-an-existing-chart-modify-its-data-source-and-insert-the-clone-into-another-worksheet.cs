// Title: Clone a Chart, Change Its Data Source, and Insert into Another Worksheet – Aspose.Cells for .NET (C#)
// Description: This Aspose.Cells for .NET example shows how to create a workbook, add a source worksheet with a column chart, then clone that chart on a different worksheet, assign a new series range, and save the file as an Excel workbook. The code demonstrates chart type copying, size preservation, and data source modification using the Aspose.Cells API.
// Keywords: Aspose.Cells chart clone C# | copy chart to another worksheet .NET | modify chart series range Aspose.Cells | Excel chart duplication Aspose.Cells | C# Aspose.Cells chart example | chart template reuse Aspose.Cells | GitHub Aspose.Cells chart sample | Aspose.Cells API chart operations
// Common Searches: how to duplicate a chart in Aspose.Cells for .NET | clone Excel chart and change data source using C# | Aspose.Cells copy chart to different worksheet | change series range after chart copy Aspose.Cells | Aspose.Cells chart cloning example GitHub
// Developer Intent: The developer needs to replicate an existing chart, point it to a new data range, and place the replicated chart on a separate worksheet using Aspose.Cells for .NET.
// Use Cases: Create a standard chart layout once and reuse it across multiple sheets with sheet‑specific data. | Generate comparative dashboards by cloning a base chart for each department or time period. | Automate monthly reporting where each month’s sheet receives a cloned chart linked to that month’s values.
// AI Prompts: Write C# code with Aspose.Cells that clones a chart from one worksheet, updates the series to a different column, and adds the clone to another worksheet. | Explain how to keep chart formatting, titles, and axis settings intact when copying a chart with Aspose.Cells. | Provide a loop example that clones a template chart for several worksheets, assigning each clone a unique data range.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    // This Aspose.Cells for .NET example shows how to create a workbook, add a source worksheet with a column chart, then clone that chart on a different worksheet, assign a new series range, and save the file as an Excel workbook. The code demonstrates chart type copying, size preservation, and data source modification using the Aspose.Cells API.
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
                srcSheet.Name = "Source";

                // Populate sample data for the chart
                srcSheet.Cells["A1"].PutValue("Category");
                srcSheet.Cells["B1"].PutValue("Value");
                srcSheet.Cells["A2"].PutValue("A");
                srcSheet.Cells["A3"].PutValue("B");
                srcSheet.Cells["A4"].PutValue("C");
                srcSheet.Cells["A5"].PutValue("D");
                srcSheet.Cells["B2"].PutValue(10);
                srcSheet.Cells["B3"].PutValue(20);
                srcSheet.Cells["B4"].PutValue(30);
                srcSheet.Cells["B5"].PutValue(40);

                // Add a chart to the source sheet
                int srcChartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart srcChart = srcSheet.Charts[srcChartIdx];
                srcChart.NSeries.Add("B2:B5", true); // Add series values
                srcChart.Title.Text = "Original Chart";

                // ---------- Destination worksheet ----------
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Populate data that will be used for the cloned chart (different source)
                destSheet.Cells["A1"].PutValue("Category");
                destSheet.Cells["C1"].PutValue("NewValue");
                destSheet.Cells["A2"].PutValue("A");
                destSheet.Cells["A3"].PutValue("B");
                destSheet.Cells["A4"].PutValue("C");
                destSheet.Cells["A5"].PutValue("D");
                destSheet.Cells["C2"].PutValue(15);
                destSheet.Cells["C3"].PutValue(25);
                destSheet.Cells["C4"].PutValue(35);
                destSheet.Cells["C5"].PutValue(45);

                // Clone the chart: create a new chart with the same type and size
                int clonedChartIdx = destSheet.Charts.Add(srcChart.Type, 5, 0, 15, 5);
                Chart clonedChart = destSheet.Charts[clonedChartIdx];

                // Add a series to the cloned chart using the new data range
                clonedChart.NSeries.Add("C2:C5", true); // Add series values
                clonedChart.Title.Text = "Cloned Chart with Modified Data Source";

                // Save the workbook (ensure the directory exists)
                string outputPath = "ChartCloneDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
