using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains a chart
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one chart
        if (sheet.Charts.Count > 0)
        {
            // Access the first chart on the sheet
            Chart chart = sheet.Charts[0];

            // Access the chart's legend
            Legend legend = chart.Legend;

            // Read the current legend position
            LegendPositionType currentPosition = legend.Position;

            // Determine the opposite corner position
            LegendPositionType newPosition;
            switch (currentPosition)
            {
                case LegendPositionType.Right:
                    newPosition = LegendPositionType.Left;
                    break;
                case LegendPositionType.Left:
                    newPosition = LegendPositionType.Right;
                    break;
                case LegendPositionType.Top:
                    newPosition = LegendPositionType.Bottom;
                    break;
                case LegendPositionType.Bottom:
                    newPosition = LegendPositionType.Top;
                    break;
                default:
                    // For other positions (e.g., Corner, NotDocked) default to Right
                    newPosition = LegendPositionType.Right;
                    break;
            }

            // Apply the new legend position
            legend.Position = newPosition;
        }

        // Save the workbook with the updated legend position
        workbook.Save("output.xlsx");
    }
}