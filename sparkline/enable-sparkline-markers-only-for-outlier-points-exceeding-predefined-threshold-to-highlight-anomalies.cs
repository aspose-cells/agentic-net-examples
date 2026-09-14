// Title: Generate a line sparkline with threshold‑based high‑ and low‑point markers and custom colors using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to add a line sparkline for a data range and automatically marks values greater than 10 in red and values less than -10 in blue. | Show how to configure a SparklineGroup in Aspose.Cells to enable ShowMarkers, ShowHighPoint, ShowLowPoint, and assign custom marker colors based on a user‑defined threshold. | Create a console application that builds a workbook, inserts sample data, adds a sparkline, and saves the file with outlier markers highlighted.
// Common Searches: Aspose.Cells C# sparkline show high point only for values above a threshold | How to set custom colors for low‑point markers in an Excel sparkline using .NET | Enable outlier markers in a line sparkline with Aspose.Cells and save as .xlsx | C# example for threshold‑driven sparkline markers in Aspose.Cells | Mark anomalous data points in Excel sparkline programmatically with Aspose.Cells
// Tags: Aspose.Cells line sparkline outlier markers | threshold based high point low point colors .NET | configure sparkline group marker visibility | custom marker colors for Excel sparkline using C# | programmatic sparkline creation Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, inserting sample data, adding a line sparkline for range A1:A8, enabling markers, and using high‑point and low‑point markers with red and blue colors to highlight values beyond a ±10 threshold, then saving the workbook as SparklineOutlierMarkersDemo.xlsx.
    public class SparklineOutlierMarkersDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data with potential outliers
                double[] data = { 5, 2, 1, 3, 25, 4, -30, 6 };
                for (int i = 0; i < data.Length; i++)
                {
                    sheet.Cells[i, 0].PutValue(data[i]); // Column A
                }

                // Define the location where the sparkline will be placed (column B, row 1)
                CellArea sparklineLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 1,
                    EndColumn = 1
                };

                // Add a sparkline group for the data range A1:A8
                int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A8", false, sparklineLocation);
                SparklineGroup group = sheet.SparklineGroups[groupIdx];

                // Add the sparkline to the group (the Add method also creates the sparkline)
                group.Sparklines.Add($"{sheet.Name}!A1:A8", 0, 1);

                // Define an outlier threshold
                double threshold = 10.0;

                // Determine if there are high or low outliers
                bool hasHighOutlier = false;
                bool hasLowOutlier = false;
                foreach (double val in data)
                {
                    if (val > threshold) hasHighOutlier = true;
                    if (val < -threshold) hasLowOutlier = true;
                }

                // Enable markers (they will appear for all points)
                group.ShowMarkers = true;
                CellsColor markersColor = workbook.CreateCellsColor();
                markersColor.Color = Color.Black;
                group.MarkersColor = markersColor;

                // Highlight high outliers using the built‑in HighPoint marker
                if (hasHighOutlier)
                {
                    group.ShowHighPoint = true;
                    CellsColor highColor = workbook.CreateCellsColor();
                    highColor.Color = Color.Red;
                    group.HighPointColor = highColor;
                }

                // Highlight low outliers using the built‑in LowPoint marker
                if (hasLowOutlier)
                {
                    group.ShowLowPoint = true;
                    CellsColor lowColor = workbook.CreateCellsColor();
                    lowColor.Color = Color.Blue;
                    group.LowPointColor = lowColor;
                }

                // Save the workbook to the current directory
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SparklineOutlierMarkersDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SparklineOutlierMarkersDemo.Run();
        }
    }
}
