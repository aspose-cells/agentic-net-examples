// Title: Create a line sparkline group in Aspose.Cells for .NET and display both high‑point and low‑point markers with custom colors
// AI Prompts: Write C# code that adds a line sparkline group to a worksheet, enables the high‑point and low‑point markers, and assigns distinct Colors to each using Aspose.Cells. | Show how to use CellsColor objects to set a green high‑point marker and a red low‑point marker for a SparklineGroup, then save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# enable high point and low point markers for a sparkline group | set custom colors for sparkline extreme points using Aspose.Cells .NET | create line sparkline with both high and low markers in Excel via Aspose.Cells | how to change sparkline high point marker color in Aspose.Cells C#
// Tags: Aspose.Cells line sparkline high low markers | set sparkline high point color Aspose.Cells | configure sparkline low point marker .NET | sparkline group extreme point customization C# | customize sparkline marker colors Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates using Aspose.Cells for .NET to create a workbook, add a line sparkline group for range A1:D1, enable both high‑point and low‑point markers, assign green to the high‑point and red to the low‑point via CellsColor, and save the file as SparklineHighLowMarkers.xlsx.
class SparklineHighLowMarkersDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4,
                StartRow = 0,    // Row 1
                EndRow = 0
            };

            // Add a sparkline group that uses the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // (Optional) Add a sparkline explicitly; the Add method already creates one
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Enable markers for both the highest and lowest points
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            // Set colors for the high and low point markers
            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            // Save the workbook with the configured sparkline
            workbook.Save("SparklineHighLowMarkers.xlsx");
            Console.WriteLine("Workbook saved successfully as SparklineHighLowMarkers.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        SparklineHighLowMarkersDemo.Run();
    }
}
