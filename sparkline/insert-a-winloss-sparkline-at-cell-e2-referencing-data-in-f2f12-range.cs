// Title: Insert a Win/Loss sparkline in cell E2 from data range F2:F12 using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, fills F2:F12 with alternating 1 and -1 values, and adds a Win/Loss sparkline to cell E2. | Show how to adjust the line weight of a Win/Loss sparkline group after it has been added to a worksheet in Aspose.Cells for .NET. | Demonstrate creating a SparklineGroup that references a specific range and places the sparkline at a single-cell location using Aspose.Cells.
// Common Searches: Aspose.Cells C# add win loss sparkline to a single cell | how to set sparkline source range F2:F12 in Aspose.Cells | customize win loss sparkline line weight Aspose.Cells .NET example | populate alternating positive and negative values for sparkline using Aspose.Cells | save workbook with sparkline as xlsx using Aspose.Cells
// Tags: win loss sparkline insertion Aspose.Cells | sparkline data range F2:F12 Aspose.Cells | set sparkline line weight .NET | populate alternating values for sparkline Aspose.Cells | save workbook with sparkline xlsx Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Creates a workbook, fills F2:F12 with 1/-1 values, adds a Win/Loss sparkline at E2, sets its line weight, and saves the file as WinLossSparkline.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the win/loss sparkline (optional)
            // F2:F12 will contain alternating positive and negative values
            for (int i = 0; i < 11; i++)
            {
                // Row index 1 = row 2, column index 5 = column F
                sheet.Cells[1 + i, 5].PutValue((i % 2 == 0) ? 1 : -1);
            }

            // Define the location cell (E2) where the sparkline will be placed
            CellArea location = CellArea.CreateCellArea("E2", "E2");

            // Add a Win/Loss sparkline group with the data range F2:F12
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.WinLoss, "F2:F12", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Optional: customize the sparkline appearance
            group.LineWeight = 1.0;

            // Save the workbook
            workbook.Save("WinLossSparkline.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
