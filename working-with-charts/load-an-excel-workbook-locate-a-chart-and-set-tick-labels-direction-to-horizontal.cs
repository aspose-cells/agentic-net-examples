using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Set the tick‑label direction of the category (X) axis to horizontal
            chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;

            // Optionally, also set the value (Y) axis tick‑label direction to horizontal
            chart.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}