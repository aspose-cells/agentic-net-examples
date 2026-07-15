using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLeaderLineDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Access the first series of the chart
            Series series = chart.NSeries[0];

            // Enable leader lines for the series
            series.HasLeaderLines = true;

            // Configure leader line appearance
            series.LeaderLines.IsAuto = false;                     // Disable automatic formatting
            series.LeaderLines.Style = LineType.Dot;               // Set line style to dotted
            series.LeaderLines.WeightPt = 1.5;                     // Set line weight (points)
            series.LeaderLines.Color = Color.Blue;                 // Set line color

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");

            Console.WriteLine("Leader line settings updated and workbook saved.");
        }
    }
}