// Title: Clone an Excel chart, change its data source, and place it on another worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clone chart and set new data range on a different worksheet | How to duplicate a column chart and modify its series source with Aspose.Cells | Programmatically copy an Excel chart to another sheet and update categories using Aspose.Cells | C# example for cloning a chart and adjusting its position and size in Aspose.Cells | Insert a cloned chart with custom NSeries ranges in Aspose.Cells workbook
// Tags: clone chart Aspose.Cells C# | modify chart data source Aspose.Cells | copy chart between worksheets Aspose.Cells | set chart NSeries range Aspose.Cells | adjust chart position size Aspose.Cells | column chart cloning Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartCloneExample
{
    // The example creates a workbook with source data and a column chart, then adds a second worksheet with its own data. It clones the chart by copying type, style, and title, clears the original series, and assigns new NSeries ranges that reference the destination sheet. The cloned chart's position and dimensions are set before saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting workbook will be saved
            string outputPath = "ClonedChartDemo.xlsx";

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source worksheet with data and a chart
                // -------------------------------------------------
                Worksheet sourceWs = workbook.Worksheets[0];
                sourceWs.Name = "SourceData";

                // Fill sample data in source worksheet (A1:B4)
                sourceWs.Cells["A1"].PutValue("Category");
                sourceWs.Cells["A2"].PutValue("A");
                sourceWs.Cells["A3"].PutValue("B");
                sourceWs.Cells["A4"].PutValue("C");
                sourceWs.Cells["B1"].PutValue("Value");
                sourceWs.Cells["B2"].PutValue(10);
                sourceWs.Cells["B3"].PutValue(20);
                sourceWs.Cells["B4"].PutValue(30);

                // Add a column chart to the source worksheet
                int sourceChartIndex = sourceWs.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart sourceChart = sourceWs.Charts[sourceChartIndex];

                // Set the data range for the source chart
                sourceChart.NSeries.Add("B2:B4", true);          // Values
                sourceChart.NSeries.CategoryData = "A2:A4";     // Categories
                sourceChart.Title.Text = "Source Chart";

                // -------------------------------------------------
                // 2. Prepare destination worksheet with its own data
                // -------------------------------------------------
                Worksheet destWs = workbook.Worksheets.Add("DestinationData");

                // Fill different sample data in destination worksheet (A1:B4)
                destWs.Cells["A1"].PutValue("Category");
                destWs.Cells["A2"].PutValue("X");
                destWs.Cells["A3"].PutValue("Y");
                destWs.Cells["A4"].PutValue("Z");
                destWs.Cells["B1"].PutValue("Value");
                destWs.Cells["B2"].PutValue(40);
                destWs.Cells["B3"].PutValue(50);
                destWs.Cells["B4"].PutValue(60);

                // -------------------------------------------------
                // 3. Clone the chart from source worksheet to destination worksheet
                // -------------------------------------------------
                // Add an empty chart placeholder in the destination worksheet
                int destChartIndex = destWs.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart destChart = destWs.Charts[destChartIndex];

                // Copy basic chart properties (type, style, etc.) from source to destination
                destChart.Type = sourceChart.Type;
                destChart.Style = sourceChart.Style;
                destChart.Title.Text = sourceChart.Title.Text;

                // -------------------------------------------------
                // 4. Modify the cloned chart's data source to point to destination data
                // -------------------------------------------------
                // Clear existing series and set new data range
                destChart.NSeries.Clear();
                destChart.NSeries.Add("B2:B4", true);          // New values from destination sheet
                destChart.NSeries.CategoryData = "A2:A4";     // New categories from destination sheet
                destChart.Title.Text = "Cloned Chart (Modified Data)";

                // Optionally adjust the chart object's position/size
                destChart.ChartObject.X = 100;
                destChart.ChartObject.Y = 200;
                destChart.ChartObject.Width = 400;
                destChart.ChartObject.Height = 300;

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
