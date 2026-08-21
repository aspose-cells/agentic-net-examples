// Title: Add a dynamic line sparkline with SparklineCollection.Add (range formula) in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill A1:A10, add a line sparkline group, insert a sparkline in B1 that references the range via SparklineCollection.Add, modify the source values, recalculate formulas, and save the file so the sparkline updates automatically.
// Keywords: Aspose.Cells | C# | .NET | SparklineCollection.Add | dynamic sparkline | range formula | auto‑refresh sparkline | CalculateFormula | Excel visualization
// Common Searches: Aspose.Cells add sparkline that updates with data | SparklineCollection.Add range string example C# | Refresh sparklines after changing source cells Aspose | Create line sparkline programmatically Aspose.Cells
// Developer Intent: Insert a line sparkline linked to a cell range that reflects data changes without manual refresh.
// Use Cases: Show a compact trend line for values in A1:A10 that stays current as the numbers are edited. | Generate several independent sparklines by calling SparklineCollection.Add with different range strings. | Ensure visualizations are up‑to‑date by invoking workbook.CalculateFormula before exporting the workbook.
// AI Prompts: Provide C# code to add a line sparkline using SparklineCollection.Add with a range string and have it update automatically. | Explain how to recalculate formulas and refresh sparklines in an Aspose.Cells workbook after modifying source data. | Show how to create multiple dynamic sparklines for various data series in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, fill A1:A10, add a line sparkline group, insert a sparkline in B1 that references the range via SparklineCollection.Add, modify the source values, recalculate formulas, and save the file so the sparkline updates automatically.
class SparklineDynamicDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1); // Values 1..10
        }

        // Add a sparkline group of type Line
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline that uses the range A1:A10 and places the sparkline in cell B1
        // SparklineCollection.Add(string dataRange, int row, int column)
        group.Sparklines.Add("A1:A10", 0, 1); // Row 0 (A1), Column 1 (B)

        // At this point the sparkline reflects the data in A1:A10.
        // Modify the source data to demonstrate automatic update.
        for (int i = 0; i < 10; i++)
        {
            // Multiply each value by 2
            double current = sheet.Cells[i, 0].DoubleValue;
            sheet.Cells[i, 0].PutValue(current * 2);
        }

        // Recalculate formulas (not strictly required for sparklines, but ensures workbook is up‑to‑date)
        workbook.CalculateFormula();

        // Save the workbook – the sparkline in B1 will display the updated data
        workbook.Save("DynamicSparklineDemo.xlsx");
    }
}
