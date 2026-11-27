// Title: Generate an Excel sparkline with Aspose.Cells for .NET while ignoring null values in the source range
// Common Searches: how to ignore empty cells when creating a sparkline with Aspose.Cells in C# | Aspose.Cells C# example for PlotEmptyCellsType.NotPlotted in sparkline groups | create line sparkline from range that contains nulls using Aspose.Cells for .NET | skip null values in Excel sparkline source data with Aspose.Cells API | prevent rendering errors caused by blank cells in Aspose.Cells sparkline
// Tags: Aspose.Cells sparkline null handling C# | configure PlotEmptyCellsType for sparkline groups | skip blank cells in Excel sparkline via Aspose API | line sparkline visual customization Aspose.Cells | create sparkline from range with missing values C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills cells A1‑E1 with numbers and intentional nulls, adds a line sparkline group referencing that range, sets PlotEmptyCellsType to NotPlotted so empty cells are ignored, applies high/low point markers and line weight, and saves the file as SparklineWithNulls.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the source range with numeric values and intentional empty cells
            // A1:E1 will be the data range for the sparkline
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(null);   // empty cell
            sheet.Cells["C1"].PutValue(30);
            sheet.Cells["D1"].PutValue(null);   // empty cell
            sheet.Cells["E1"].PutValue(50);

            // Define the destination cell area for the sparkline (column F, same row)
            CellArea sparkArea = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 5, // Column F (0‑based index)
                EndColumn = 5
            };

            // Add a sparkline group; the source range includes empty cells
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,   // Sparkline type
                "A1:E1",              // Source data range
                false,                // Do not plot right‑to‑left
                sparkArea);           // Destination area

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Skip empty cells (do not plot them)
            group.PlotEmptyCellsType = PlotEmptyCellsType.NotPlotted;

            // Optional visual customizations
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;
            group.HighPointColor.Color = Color.Green;
            group.LowPointColor.Color = Color.Red;
            group.LineWeight = 1.0;

            // Save the workbook
            string outputPath = "SparklineWithNulls.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
