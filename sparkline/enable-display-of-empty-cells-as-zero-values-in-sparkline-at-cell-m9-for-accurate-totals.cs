// Title: How to configure a line sparkline to treat empty cells as zeros in Aspose.Cells for .NET (C#) and place it in cell M9
// Common Searches: Aspose.Cells C# set sparkline empty cells to zero in Excel worksheet | Create line sparkline at M9 that plots blanks as zero using Aspose.Cells | PlotEmptyCellsType.Zero example for sparkline in Aspose.Cells .NET | Add sparkline to a specific cell and display missing values as zero in C# | Aspose.Cells sparkline configuration for handling empty data points
// Tags: Aspose.Cells line sparkline empty cells zero | C# PlotEmptyCellsType Zero sparkline | Aspose.Cells add sparkline to cell M9 | Excel sparkline blank values handling Aspose.Cells | Aspose.Cells sparkline series color blue

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills alternating cells in row 9, adds a line sparkline sourced from A9:L9 into cell M9, sets PlotEmptyCellsType to Zero so empty cells are displayed as zero, applies a blue series color, and saves the file as SparklineEmptyAsZero.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate row 9 (index 8) with sample data.
            // Some cells are left empty to demonstrate the "zero for empty" behavior.
            for (int col = 0; col <= 11; col++) // Columns A (0) to L (11)
            {
                if (col % 2 == 0) // Fill every other cell
                {
                    sheet.Cells[8, col].PutValue(col + 1); // Example value
                }
                // Odd columns remain empty
            }

            // Define the source range for the sparkline (A9:L9)
            string sourceRange = "A9:L9";

            // Define the location cell for the sparkline (M9)
            string locationCell = "M9";

            // Add a line sparkline group.
            // The first parameter (0) is the sparkline type code for a line sparkline.
            int groupIndex = sheet.SparklineGroups.Add(
                0,                     // Line sparkline type
                sourceRange,           // Data source range
                false,                 // No data markers
                CellArea.CreateCellArea(locationCell, locationCell) // Destination cell area
            );

            // Retrieve the created sparkline group
            SparklineGroup sparkGroup = sheet.SparklineGroups[groupIndex];

            // Configure the group to plot empty cells as zero values
            sparkGroup.PlotEmptyCellsType = PlotEmptyCellsType.Zero;

            // Set a visible color for the sparkline (using CellsColor)
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Blue;
            sparkGroup.SeriesColor = seriesColor;

            // Save the workbook
            string outputPath = "SparklineEmptyAsZero.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            Console.WriteLine("The sparkline in cell M9 now displays empty cells as zero.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
