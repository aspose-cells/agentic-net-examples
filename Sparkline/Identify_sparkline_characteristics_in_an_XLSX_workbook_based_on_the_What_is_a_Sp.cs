using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class SparklineInfoExtractor
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Get the collection of sparkline groups on the current worksheet
            SparklineGroupCollection groups = sheet.SparklineGroups;

            if (groups.Count == 0)
            {
                Console.WriteLine("  No sparkline groups found.");
                continue;
            }

            // Process each sparkline group
            for (int i = 0; i < groups.Count; i++)
            {
                SparklineGroup group = groups[i];
                Console.WriteLine($"  Sparkline Group {i}:");
                Console.WriteLine($"    Type: {group.Type}");
                Console.WriteLine($"    DisplayHidden: {group.DisplayHidden}");
                Console.WriteLine($"    ShowHighPoint: {group.ShowHighPoint}");
                Console.WriteLine($"    ShowLowPoint: {group.ShowLowPoint}");
                Console.WriteLine($"    SeriesColor: {group.SeriesColor?.Color}");
                Console.WriteLine($"    HighPointColor: {group.HighPointColor?.Color}");
                Console.WriteLine($"    LowPointColor: {group.LowPointColor?.Color}");
                Console.WriteLine($"    PresetStyle: {group.PresetStyle}");
                Console.WriteLine($"    LineWeight: {group.LineWeight}");

                // Iterate through each sparkline within the group
                SparklineCollection sparklines = group.Sparklines;
                for (int j = 0; j < sparklines.Count; j++)
                {
                    Sparkline sp = sparklines[j];
                    Console.WriteLine($"    Sparkline {j}:");
                    Console.WriteLine($"      Row: {sp.Row}, Column: {sp.Column}");
                    Console.WriteLine($"      DataRange: {sp.DataRange}");
                }
            }
        }

        // Save the workbook (no modifications made, but required by lifecycle rules)
        workbook.Save("output.xlsx");
    }
}