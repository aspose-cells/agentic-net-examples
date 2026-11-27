// Title: How to clone a chart shape, modify its data series, and place the cloned chart on another worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# clone chart shape and bind to a different data range on another sheet | Copy chart formatting and position to a new worksheet with Aspose.Cells .NET | Duplicate an Excel column chart and change its series values using Aspose.Cells in C# | Transfer a chart object from one worksheet to another while preserving size and placement | Programmatically create a cloned chart with new data in an Aspose.Cells workbook
// Tags: chart shape cloning Aspose.Cells | modify chart series Aspose.Cells C# | copy chart formatting Aspose.Cells | place chart on another worksheet Aspose.Cells | clone chart with new data range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace CloneChartExample
{
    // The example creates a source worksheet with sample data and a column chart, retrieves the chart's shape, adds a destination worksheet with different data, creates a new chart of the same type, clears its default series, binds it to the new data range, copies visual properties (placement, size, name) from the source chart shape, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source worksheet with data and a chart
                // -------------------------------------------------
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "Source";

                // Sample data for the source chart
                srcSheet.Cells["A1"].PutValue("Category");
                srcSheet.Cells["A2"].PutValue("A");
                srcSheet.Cells["A3"].PutValue("B");
                srcSheet.Cells["A4"].PutValue("C");

                srcSheet.Cells["B1"].PutValue("Value");
                srcSheet.Cells["B2"].PutValue(10);
                srcSheet.Cells["B3"].PutValue(20);
                srcSheet.Cells["B4"].PutValue(30);

                // Add a column chart to the source sheet
                int srcChartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart srcChart = srcSheet.Charts[srcChartIdx];

                // Set the data range for the source chart
                srcChart.NSeries.Add("B2:B4", true);
                srcChart.NSeries.CategoryData = "A2:A4";
                srcChart.Title.Text = "Source Chart";

                // -------------------------------------------------
                // 2. Get the source chart shape (visual representation)
                // -------------------------------------------------
                ChartShape srcChartShape = srcChart.ChartObject;

                // -------------------------------------------------
                // 3. Prepare destination worksheet with its own data
                // -------------------------------------------------
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Sample data for the destination chart (different values)
                destSheet.Cells["A1"].PutValue("Category");
                destSheet.Cells["A2"].PutValue("X");
                destSheet.Cells["A3"].PutValue("Y");
                destSheet.Cells["A4"].PutValue("Z");

                destSheet.Cells["B1"].PutValue("Value");
                destSheet.Cells["B2"].PutValue(40);
                destSheet.Cells["B3"].PutValue(50);
                destSheet.Cells["B4"].PutValue(60);

                // -------------------------------------------------
                // 4. Add a new chart to the destination sheet
                // -------------------------------------------------
                // Use placeholder positions; visual placement will be set later
                int destChartIdx = destSheet.Charts.Add(srcChart.Type, 0, 0, 10, 5);
                Chart destChart = destSheet.Charts[destChartIdx];

                // Clear any automatically created series and bind to destination data
                destChart.NSeries.Clear();
                destChart.NSeries.Add("B2:B4", true);
                destChart.NSeries.CategoryData = "A2:A4";
                destChart.Title.Text = "Cloned Chart (Modified Data)";

                // -------------------------------------------------
                // 5. Copy visual formatting from the source chart to the destination chart
                // -------------------------------------------------
                destChart.ChartObject.Placement = srcChartShape.Placement;
                destChart.ChartObject.X = srcChartShape.X;
                destChart.ChartObject.Y = srcChartShape.Y;
                destChart.ChartObject.Width = srcChartShape.Width;
                destChart.ChartObject.Height = srcChartShape.Height;
                destChart.ChartObject.Name = "ClonedChartShape";

                // -------------------------------------------------
                // 6. Save the workbook
                // -------------------------------------------------
                string outFile = "ClonedChartExample.xlsx";

                // Ensure the directory exists (in case a path is provided)
                string outDir = Path.GetDirectoryName(outFile);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
