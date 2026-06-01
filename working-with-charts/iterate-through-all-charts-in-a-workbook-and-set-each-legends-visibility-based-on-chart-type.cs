using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through all charts on the current worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Set legend visibility based on chart type.
                // Hide legends for pie‑type charts (they are often self‑explanatory),
                // show legends for all other chart types.
                switch (chart.Type)
                {
                    case ChartType.Pie:
                    case ChartType.Pie3D:
                    case ChartType.PieExploded:
                    case ChartType.Pie3DExploded:
                    case ChartType.Doughnut:
                    case ChartType.DoughnutExploded:
                        chart.ShowLegend = false;
                        break;

                    default:
                        chart.ShowLegend = true;
                        break;
                }
            }
        }

        // Save the workbook with updated legend settings
        workbook.Save("output.xlsx");
    }
}