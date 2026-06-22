using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source worksheet with sample data
                // -------------------------------------------------
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceSheet";

                // Populate data for the chart
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["A2"].PutValue("A");
                sourceSheet.Cells["A3"].PutValue("B");
                sourceSheet.Cells["A4"].PutValue("C");

                sourceSheet.Cells["B1"].PutValue("Value");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["B3"].PutValue(20);
                sourceSheet.Cells["B4"].PutValue(30);

                // -------------------------------------------------
                // 2. Add an original chart to the source worksheet
                // -------------------------------------------------
                int originalChartIndex = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart originalChart = sourceSheet.Charts[originalChartIndex];
                originalChart.SetChartDataRange("A1:B4", true);
                originalChart.Title.Text = "Original Chart";

                // -------------------------------------------------
                // 3. Create a destination worksheet where the clone will be placed
                // -------------------------------------------------
                Worksheet destSheet = workbook.Worksheets.Add("ClonedSheet");

                // -------------------------------------------------
                // 4. Clone the chart
                // -------------------------------------------------
                // Add a new chart to the destination sheet with the same type and approximate size
                int clonedChartIndex = destSheet.Charts.Add(
                    originalChart.Type,
                    originalChart.ChartObject.UpperLeftRow,
                    originalChart.ChartObject.UpperLeftColumn,
                    originalChart.ChartObject.UpperLeftRow + 10,   // lower‑right row (approx.)
                    originalChart.ChartObject.UpperLeftColumn + 5 // lower‑right column (approx.)
                );
                Chart clonedChart = destSheet.Charts[clonedChartIndex];

                // Copy the data range from the original chart
                string dataRange = originalChart.GetChartDataRange();
                clonedChart.SetChartDataRange(dataRange, true);

                // Copy the title and modify it
                clonedChart.Title.Text = originalChart.Title.Text + " (Clone)";

                // Copy visual properties
                clonedChart.Style = originalChart.Style;
                clonedChart.ShowLegend = originalChart.ShowLegend;

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                string outputPath = "ChartCloneDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}