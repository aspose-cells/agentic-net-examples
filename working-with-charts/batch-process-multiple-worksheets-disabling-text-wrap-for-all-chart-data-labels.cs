using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DisableChartDataLabelWrap
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts on the current worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Iterate through each series in the chart
                foreach (Series series in chart.NSeries)
                {
                    // Disable text wrap for the series data labels
                    series.DataLabels.IsTextWrapped = false;

                    // Also disable text wrap for data labels of individual points, if any
                    foreach (ChartPoint point in series.Points)
                    {
                        point.DataLabels.IsTextWrapped = false;
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}