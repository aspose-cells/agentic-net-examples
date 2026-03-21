using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineReport
{
    static void Main()
    {
        // Load the workbook from a file (replace with actual path if needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Access the collection of sparkline groups in the current worksheet
            SparklineGroupCollection groups = sheet.SparklineGroups;

            // Process each sparkline group
            for (int i = 0; i < groups.Count; i++)
            {
                SparklineGroup group = groups[i];
                Console.WriteLine($"  Sparkline Group {i}:");
                Console.WriteLine($"    Type: {group.Type}");
                Console.WriteLine($"    DisplayHidden: {group.DisplayHidden}");
                Console.WriteLine($"    ShowHighPoint: {group.ShowHighPoint}");
                Console.WriteLine($"    ShowLowPoint: {group.ShowLowPoint}");
                Console.WriteLine($"    ShowFirstPoint: {group.ShowFirstPoint}");
                Console.WriteLine($"    ShowLastPoint: {group.ShowLastPoint}");
                Console.WriteLine($"    ShowMarkers: {group.ShowMarkers}");
                Console.WriteLine($"    ShowNegativePoints: {group.ShowNegativePoints}");
                Console.WriteLine($"    PlotEmptyCellsType: {group.PlotEmptyCellsType}");
                Console.WriteLine($"    PlotRightToLeft: {group.PlotRightToLeft}");
                Console.WriteLine($"    SeriesColor: {group.SeriesColor?.Color}");
                Console.WriteLine($"    HighPointColor: {group.HighPointColor?.Color}");
                Console.WriteLine($"    LowPointColor: {group.LowPointColor?.Color}");
                Console.WriteLine($"    FirstPointColor: {group.FirstPointColor?.Color}");
                Console.WriteLine($"    LastPointColor: {group.LastPointColor?.Color}");
                Console.WriteLine($"    MarkersColor: {group.MarkersColor?.Color}");
                Console.WriteLine($"    NegativePointsColor: {group.NegativePointsColor?.Color}");
                Console.WriteLine($"    LineWeight: {group.LineWeight}");

                // Iterate through each sparkline within the group
                for (int j = 0; j < group.Sparklines.Count; j++)
                {
                    Sparkline spark = group.Sparklines[j];
                    Console.WriteLine($"    Sparkline {j}: Row={spark.Row}, Column={spark.Column}, DataRange={spark.DataRange}");
                }
            }
        }

        // Save a copy of the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("output_report.xlsx");
    }
}