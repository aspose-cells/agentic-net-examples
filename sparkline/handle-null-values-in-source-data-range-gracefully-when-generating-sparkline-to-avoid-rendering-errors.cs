// Title: Create a line sparkline in Aspose.Cells for .NET while skipping empty cells in the source range
// Common Searches: Aspose.Cells C# create line sparkline that ignores blank cells in source range | how to prevent sparkline rendering errors caused by null values using Aspose.Cells | C# Aspose.Cells sparkline group hide markers for empty data points | generate sparkline from range with missing values in Aspose.Cells .NET | skip null cells when adding sparkline to worksheet with Aspose.Cells
// Tags: line sparkline creation Aspose.Cells C# | null cell handling in sparkline source range | SparklineGroup ShowMarkers property usage | CellArea definition for sparkline data range | sparkline rendering with missing data points

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills column A with numbers leaving specific cells blank to simulate nulls, defines the source range A1:A7 using CellArea, adds a line sparkline to cell B1, disables markers for empty cells via ShowMarkers = false, and saves the file as SparklineWithNulls.xlsx, ensuring null values do not cause rendering errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data range A1:A7 with numbers, leaving some cells empty (null)
            double[] data = { 10, 20, 30, 40, 50, 60, 70 };
            // Indices 2 and 5 will be left empty to simulate nulls
            for (int i = 0; i < data.Length; i++)
            {
                if (i == 2 || i == 5) // skip these indices to create empty cells
                    continue;

                sheet.Cells[i, 0].PutValue(data[i]); // Column A (index 0)
            }

            // Define source data range as CellArea
            CellArea sourceRange = CellArea.CreateCellArea("A1", "A7");

            // Add a sparkline group of type Line in cell B1 referencing the source range A1:A7
            // The Add method returns the index of the newly created group
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "B1", false, sourceRange);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

            // Optional: hide markers for empty cells to keep the visual clean
            sparklineGroup.ShowMarkers = false;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("SparklineWithNulls.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
