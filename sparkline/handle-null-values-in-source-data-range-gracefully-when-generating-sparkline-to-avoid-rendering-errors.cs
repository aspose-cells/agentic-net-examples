// Title: Generate a line sparkline in Aspose.Cells for .NET while configuring empty‑cell handling to prevent rendering errors
// AI Prompts: Write C# code that adds a line sparkline to a worksheet using Aspose.Cells and sets the sparkline group to plot empty cells as zeros. | Show how to create a SparklineGroup in Aspose.Cells and configure it to interpolate missing values instead of skipping them. | Provide a C# example that sets PlotEmptyCellsType to NotPlotted for a sparkline range to hide null cells.
// Common Searches: Aspose.Cells how to set PlotEmptyCellsType for a sparkline in C# | C# sparkline ignore empty cells Aspose.Cells example | Configure empty cell handling for line sparkline using Aspose.Cells .NET
// Tags: Aspose.Cells line sparkline empty cells | PlotEmptyCellsType zero setting | sparkline group options Aspose.Cells | C# handling null values in sparkline | interpolate missing sparkline points .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills cells A1‑E1 with numbers and intentional empty cells, adds a line sparkline at F1 covering that range, configures PlotEmptyCellsType to define how empty cells are rendered (e.g., as zeros, interpolated, or not plotted), and saves the file as SparklineNullHandlingDemo.xlsx.
public class SparklineNullHandlingDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the source range with some values and intentional nulls (empty cells)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(null); // empty cell
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(null); // empty cell
        sheet.Cells["E1"].PutValue(3);

        // Define the location where the sparkline will be placed (column F, row 1)
        CellArea location = new CellArea
        {
            StartColumn = 5, // column index for "F"
            EndColumn = 5,
            StartRow = 0,    // row index for "1"
            EndRow = 0
        };

        // Add a sparkline group for the data range A1:E1.
        // This call also creates a sparkline automatically.
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Configure how empty cells are plotted.
        // Options: PlotEmptyCellsType.Zero, PlotEmptyCellsType.Interpolated, PlotEmptyCellsType.NotPlotted
        group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;

        // Save the workbook to a file
        string outputPath = "SparklineNullHandlingDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
