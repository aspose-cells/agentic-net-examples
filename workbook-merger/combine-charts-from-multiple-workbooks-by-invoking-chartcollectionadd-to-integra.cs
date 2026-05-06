using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace CombineChartsDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of source workbooks that contain charts
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

            // Create the destination workbook where all charts will be merged
            Workbook destWorkbook = new Workbook();                     // create
            Worksheet destSheet = destWorkbook.Worksheets[0];          // use first sheet

            // Iterate through each source workbook
            foreach (string file in sourceFiles)
            {
                // Load the source workbook
                Workbook srcWorkbook = new Workbook(file);             // load

                // For simplicity, work with the first worksheet of the source
                Worksheet srcSheet = srcWorkbook.Worksheets[0];
                ChartCollection srcCharts = srcSheet.Charts;

                // Copy every chart from the source worksheet to the destination worksheet
                for (int i = 0; i < srcCharts.Count; i++)
                {
                    Chart srcChart = srcCharts[i];

                    // Add a new chart to the destination sheet using the same chart type.
                    // Position is set arbitrarily; you can adjust as needed.
                    int chartIndex = destSheet.Charts.Add(
                        srcChart.Type,          // Chart type
                        5, 0, 25, 10);          // topRow, leftColumn, bottomRow, rightColumn

                    Chart destChart = destSheet.Charts[chartIndex];

                    // Copy series data ranges from the source chart to the new chart
                    foreach (Series srcSeries in srcChart.NSeries)
                    {
                        // Use the Values formula of the source series
                        destChart.NSeries.Add(srcSeries.Values, true);
                    }

                    // Copy category (X‑axis) data if it is defined
                    if (!string.IsNullOrEmpty(srcChart.NSeries.CategoryData))
                    {
                        destChart.NSeries.CategoryData = srcChart.NSeries.CategoryData;
                    }

                    // Copy chart title
                    if (!string.IsNullOrEmpty(srcChart.Title.Text))
                    {
                        destChart.Title.Text = srcChart.Title.Text;
                    }

                    // Optional: copy legend visibility
                    destChart.ShowLegend = srcChart.ShowLegend;
                }
            }

            // Save the combined workbook
            destWorkbook.Save("CombinedCharts.xlsx", SaveFormat.Xlsx);   // save
        }
    }
}