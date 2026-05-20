using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineOutlierDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data – values above 15 will be considered outliers
            double[] data = { 5, 8, 12, 20, 7, 30, 9, 4 };
            for (int i = 0; i < data.Length; i++)
            {
                sheet.Cells[0, i].PutValue(data[i]);
            }

            // Define the location where the sparkline will be placed (cell E1)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a line sparkline group for the data range A1:H1
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:H1", false, sparklineLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add the sparkline to the group (the Add method also creates the sparkline)
            group.Sparklines.Add(sheet.Name + "!A1:H1", 0, 4);

            // Enable markers – this will display a marker for each point
            group.ShowMarkers = true;

            // Set marker color to red to make outliers stand out
            CellsColor markerColor = workbook.CreateCellsColor();
            markerColor.Color = Color.Red;
            group.MarkersColor = markerColor;

            // OPTIONAL: Highlight the highest and lowest points as well
            group.ShowHighPoint = true;
            group.HighPointColor = workbook.CreateCellsColor();
            group.HighPointColor.Color = Color.Green;

            group.ShowLowPoint = true;
            group.LowPointColor = workbook.CreateCellsColor();
            group.LowPointColor.Color = Color.Blue;

            // Save the workbook
            workbook.Save("SparklineOutlierDemo.xlsx");
        }
    }
}