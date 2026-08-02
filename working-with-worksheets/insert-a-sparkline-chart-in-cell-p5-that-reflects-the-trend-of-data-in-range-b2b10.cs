// Title: Add a line sparkline to cell P5 from range B2:B10 using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills B2:B10 with sample numbers, inserts a line‑type sparkline in cell P5 that references the range, sets the series color to blue, and saves the file as SparklineInP5.xlsx.
// Keywords: Aspose.Cells | C# | sparkline | line sparkline | cell P5 | range B2:B10 | add sparkline | sparkline group | set series color | save workbook | Excel dashboard
// Common Searches: Aspose.Cells add sparkline to a single cell | C# sparkline from range B2:B10 | How to set sparkline color in Aspose.Cells | Save workbook with sparkline using Aspose.Cells | Create line sparkline in cell P5 .NET
// Developer Intent: Insert a line sparkline that visualizes the values in B2:B10 into cell P5.
// Use Cases: Display a compact monthly sales trend in a financial report. | Add a sparkline to each row of a KPI dashboard for quick performance insight. | Generate lightweight trend indicators in exported Excel files for stakeholders. | Customize sparkline appearance (color, markers) before sharing the workbook.
// AI Prompts: Write C# code to create a column sparkline in cell D4 from range C2:C12 with a red series color using Aspose.Cells. | Explain how to modify sparkline marker style and line weight after the sparkline group has been added. | Show an example of adding multiple sparkline groups on the same worksheet, each bound to a different data range. | Provide code to bind a sparkline to a dynamic named range in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills B2:B10 with sample numbers, inserts a line‑type sparkline in cell P5 that references the range, sets the series color to blue, and saves the file as SparklineInP5.xlsx.
class SparklineInCellP5
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in the range B2:B10 (column index 1)
            for (int row = 2; row <= 10; row++)
            {
                sheet.Cells[row, 1].PutValue(row * 5); // example values
            }

            // Define the location cell area for the sparkline (single cell P5)
            CellArea location = CellArea.CreateCellArea("P5", "P5");

            // Add a sparkline group of type Line that uses the data range B2:B10
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "B2:B10", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Optional: customize the appearance of the sparkline
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Blue;
            group.SeriesColor = seriesColor;

            // Save the workbook
            string outputPath = "SparklineInP5.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
