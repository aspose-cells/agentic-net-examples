// Title: Create line sparklines for rows 1‑20 in column U with Aspose.Cells for .NET (C#)
// Description: C# code that builds a 20 × 20 table, adds a line‑type SparklineGroup, inserts a sparkline for each row (data range A‑T) into column U, and saves the workbook as BatchSparklines.xlsx.
// Keywords: Aspose.Cells | C# | .NET | sparkline | line sparkline | batch sparklines | SparklineGroup | add sparklines programmatically | Excel automation | generate sparklines | worksheet
// Common Searches: Aspose.Cells add sparklines to each row | C# batch create sparklines in Excel | how to generate line sparklines with Aspose.Cells | programmatically add sparkline group .NET | create sparkline for multiple rows Aspose
// Developer Intent: Insert a line sparkline for every row of a 20‑row range, using that row’s A‑T values, and place the sparkline in column U.
// Use Cases: Show monthly sales trends for 20 products, with each product’s data in columns A‑T and a sparkline in column U. | Display sensor‑reading trends for 20 devices, where each row holds sequential measurements and the adjacent sparkline provides a quick visual cue. | Build a KPI dashboard that automatically adds sparklines for each category row to illustrate performance without manual charting. | Create a financial report that visualizes 20 fiscal periods, using sparklines to compare month‑over‑month changes side‑by‑side.
// AI Prompts: Generate C# code that adds column‑type sparklines for rows 1‑30 using data range B‑U and saves the file. | Modify the example to use SparklineType.Column, set custom marker colors, and enable axis display. | Explain how to change the line color, weight, and transparency of sparklines after they have been added to a worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that builds a 20 × 20 table, adds a line‑type SparklineGroup, inserts a sparkline for each row (data range A‑T) into column U, and saves the workbook as BatchSparklines.xlsx.
class BatchSparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for rows 1‑20, columns A‑T (0‑19)
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                // Example data: (row + 1) * (col + 1)
                sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
            }
        }

        // Add a sparkline group of type Line (no initial sparklines)
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // For each row, add a sparkline that uses the data range A‑T of that row
        // and places the sparkline in column U (index 20)
        for (int row = 0; row < 20; row++)
        {
            string dataRange = $"A{row + 1}:T{row + 1}";
            int sparklineRow = row;      // zero‑based row index for the sparkline location
            int sparklineColumn = 20;    // column U (zero‑based)

            group.Sparklines.Add(dataRange, sparklineRow, sparklineColumn);
        }

        // Save the workbook
        workbook.Save("BatchSparklines.xlsx");
    }
}
