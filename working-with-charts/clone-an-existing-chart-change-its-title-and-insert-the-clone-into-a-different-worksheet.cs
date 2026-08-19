// Title: Clone a Chart, Update Its Title, and Insert It into Another Worksheet with AspNet.Cells for .NET
// Description: Shows how to create a workbook, add sample data, build a column chart, copy the chart to a new worksheet, reuse its data range, change the cloned chart title, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | chart cloning | copy chart to another worksheet | change chart title | GetChartDataRange | SetChartDataRange | Excel chart duplication .NET | Aspose.Cells API example
// Common Searches: Aspose.Cells copy chart to another sheet | How to clone an Excel chart with Aspose.Cells | Change chart title after cloning Aspose.Cells | SetChartDataRange usage Aspose.Cells | C# duplicate chart programmatically | Aspose.Cells chart clone example
// Developer Intent: Copy an existing chart, modify its title, and place the copy on a different worksheet programmatically.
// Use Cases: Create a dashboard sheet that aggregates identical charts from multiple source sheets, each with a custom heading. | Generate a multi‑section report where the same chart layout is reused, only the title varies per section. | Automate the migration of a chart to a summary worksheet while preserving its data source and chart type.
// AI Prompts: Provide C# code to clone an Aspose.Cells chart, keep its data range, and set a new title on a different worksheet. | Explain how GetChartDataRange and SetChartDataRange work when moving a chart between sheets in Aspose.Cells. | Show an example of copying a column chart to another sheet without affecting the original chart in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    // Shows how to create a workbook, add sample data, build a column chart, copy the chart to a new worksheet, reuse its data range, change the cloned chart title, and save the file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Source worksheet with original chart ----------
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceSheet";

                // Populate sample data
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Value");
                sourceSheet.Cells["A2"].PutValue("A");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["A3"].PutValue("B");
                sourceSheet.Cells["B3"].PutValue(20);
                sourceSheet.Cells["A4"].PutValue("C");
                sourceSheet.Cells["B4"].PutValue(30);

                // Add a chart to the source sheet
                int srcUpperLeftRow = 5;
                int srcUpperLeftColumn = 0;
                int srcBottomRightRow = 15;
                int srcBottomRightColumn = 5;

                // Add returns the index of the newly created chart
                int srcChartIndex = sourceSheet.Charts.Add(ChartType.Column,
                                                          srcUpperLeftRow,
                                                          srcUpperLeftColumn,
                                                          srcBottomRightRow,
                                                          srcBottomRightColumn);
                Chart srcChart = sourceSheet.Charts[srcChartIndex];
                srcChart.NSeries.Add("B2:B4", true);
                srcChart.NSeries.CategoryData = "A2:A4";
                srcChart.Title.Text = "Original Chart Title";

                // ---------- Target worksheet where the chart will be cloned ----------
                Worksheet targetSheet = workbook.Worksheets.Add("TargetSheet");

                // Add a new chart with the same type and position on the target sheet
                int clonedChartIndex = targetSheet.Charts.Add(srcChart.Type,
                                                             srcUpperLeftRow,
                                                             srcUpperLeftColumn,
                                                             srcBottomRightRow,
                                                             srcBottomRightColumn);
                Chart clonedChart = targetSheet.Charts[clonedChartIndex];

                // Copy the data range from the source chart
                string dataRange = srcChart.GetChartDataRange(); // e.g., "SourceSheet!$B$2:$B$4"
                clonedChart.SetChartDataRange(dataRange, true);

                // Change the title of the cloned chart
                clonedChart.Title.Text = "Cloned Chart Title";

                // Save the workbook
                workbook.Save("ChartCloneResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
