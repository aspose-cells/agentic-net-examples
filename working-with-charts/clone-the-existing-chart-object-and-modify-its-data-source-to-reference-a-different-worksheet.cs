// Title: Clone an Excel chart and change its data source to a new worksheet using Aspose.Cells C#
// Description: Loads a workbook, extracts the first chart, adds a new worksheet, creates a chart of the same type and position, replaces the original sheet name in the chart's data range with the new sheet name, applies the updated range with SetChartDataRange, copies title, style and legend settings, and saves the file.
// Keywords: Aspose.Cells clone chart C# | change chart data source Aspose.Cells | duplicate Excel chart programmatically | SetChartDataRange Aspose.Cells | GetChartDataRange example | copy chart to another worksheet | Aspose.Cells chart manipulation
// Common Searches: clone chart to different sheet Aspose.Cells | update chart data range after copying C# | Aspose.Cells set chart data source worksheet | copy Excel chart programmatically .NET | Aspose.Cells chart cloning tutorial
// Developer Intent: Create a copy of an existing Excel chart and point its series to a different worksheet while preserving visual formatting.
// Use Cases: Generate a summary sheet that shows visual copies of charts from a data sheet, keeping source data isolated. | Automate report creation where each chart is duplicated on its own worksheet for independent formatting or printing. | Build a dashboard template that clones charts onto a dedicated sheet and redirects their data ranges to dynamically populated worksheets.
// AI Prompts: Show me C# code to clone an Excel chart and redirect its data range to another worksheet with Aspose.Cells. | Provide an example that copies a chart, updates the series source to a different sheet, and retains the original style and title. | Explain how to replace the sheet name in a chart's data range string when cloning charts using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartClone
{
    // Loads a workbook, extracts the first chart, adds a new worksheet, creates a chart of the same type and position, replaces the original sheet name in the chart's data range with the new sheet name, applies the updated range with SetChartDataRange, copies title, style and legend settings, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for input and output workbooks
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists; if not, create a simple workbook with a chart
                if (!File.Exists(inputPath))
                {
                    CreateSampleWorkbook(inputPath);
                }

                // Load the workbook that contains the source chart
                Workbook workbook = new Workbook(inputPath);

                // Assume the first worksheet contains at least one chart
                Worksheet sourceSheet = workbook.Worksheets[0];
                if (sourceSheet.Charts.Count == 0)
                {
                    throw new InvalidOperationException("Source worksheet does not contain any charts.");
                }

                Chart sourceChart = sourceSheet.Charts[0];

                // Add a new worksheet to host the cloned chart
                Worksheet targetSheet = workbook.Worksheets.Add("ClonedChartSheet");

                // Add a new chart to the target sheet using the same type and position as the source chart
                int newChartIndex = targetSheet.Charts.Add(
                    sourceChart.Type,
                    sourceChart.ChartObject.UpperLeftRow,
                    sourceChart.ChartObject.UpperLeftColumn,
                    sourceChart.ChartObject.LowerRightRow,
                    sourceChart.ChartObject.LowerRightColumn);
                Chart clonedChart = targetSheet.Charts[newChartIndex];

                // Get the data range of the source chart (e.g., "Sheet1!A1:B4")
                string sourceDataRange = sourceChart.GetChartDataRange();

                // Redirect the data range to the new worksheet
                string targetDataRange = sourceDataRange.Replace(sourceSheet.Name, targetSheet.Name);

                // Apply the modified data range to the cloned chart
                clonedChart.SetChartDataRange(targetDataRange, true);

                // Copy visual properties from the source chart
                clonedChart.Title.Text = sourceChart.Title.Text;
                clonedChart.Style = sourceChart.Style;
                clonedChart.ShowLegend = sourceChart.ShowLegend;

                // Save the workbook with the cloned chart
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a simple workbook with sample data and a chart for demonstration purposes
        private static void CreateSampleWorkbook(string path)
        {
            try
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Name = "Sheet1";

                // Populate sample data
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["A3"].PutValue("B");
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["B3"].PutValue(20);
                ws.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = ws.Charts[chartIndex];
                chart.NSeries.Add("Sheet1!B2:B4", true);
                // Category data assignment removed due to API compatibility
                chart.Title.Text = "Sample Chart";

                wb.Save(path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
            }
        }
    }
}
