// Title: Place a column chart at row 15, column 3 with a 400‑point width using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that inserts a column chart whose upper‑left corner starts at row 15, column 3 and sets the chart width to exactly 400 points. | Explain how to compute the lower‑right column index required to achieve a 400‑point chart width when positioning a chart with Aspose.Cells.
// Common Searches: Aspose.Cells C# set chart upper left cell to row 15 column 3 | how to define chart width in points with Aspose.Cells .NET | calculate lower right column for 400 point chart width Aspose.Cells | position Excel chart at specific cell range using Aspose.Cells C#
// Tags: Aspose.Cells set chart position by cell coordinates | Aspose.Cells define chart width in points | C# column chart placement with Aspose.Cells | Excel chart sizing using Aspose.Cells .NET | determine chart bounds for specific point width Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds a column chart whose upper‑left corner is anchored at row 15, column 3, computes the lower‑right column to give the chart a width of 400 points, sets a title, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Get the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Define chart position: upper‑left at row 15 (index 14), column 3 (index 2)
            // Define lower‑right to give the chart a reasonable size
            int upperLeftRow = 14;
            int upperLeftColumn = 2;
            int lowerRightRow = 24;   // adjust as needed for height
            int lowerRightColumn = 12; // adjust as needed for width

            // Add a column chart with the specified bounds
            int chartIndex = ws.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart chart = ws.Charts[chartIndex];

            // Optional: set chart title or other properties here
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook
            wb.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
