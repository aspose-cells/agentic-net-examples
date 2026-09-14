// Title: Create a line sparkline in column B from the A1:A10 range with Aspose.Cells in C#
// AI Prompts: Generate a workbook, fill cells A1 through A10 with sequential numbers, add a line sparkline group, and place the sparkline in cell B1 using Aspose.Cells C#. | Use SparklineCollection.Add to insert a line sparkline referencing A1:A10 into column B of a new worksheet and save the file. | Write C# code that populates a data range, creates a line sparkline group, and adds the sparkline to column B via Aspose.Cells.
// Common Searches: Aspose.Cells how to add a line sparkline to column B in C# | C# example for SparklineCollection.Add with range A1:A10 | Create line sparkline in Excel using Aspose.Cells C# code sample | Saving a workbook with a line sparkline using Aspose.Cells | Populate A1:A10 and generate sparkline in B1 with Aspose.Cells
// Tags: sparklinegroup line type Aspose.Cells | insert sparkline into B1 C# | excel line sparkline generation Aspose.Cells | range A1:A10 sparkline data | workbook save with sparkline

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills cells A1‑A10 with values 1‑10, adds a line sparkline group, inserts a line sparkline into cell B1 referencing that range, and saves the file as LineSparkline.xlsx.
class AddLineSparkline
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the data range A1:A10 with sample values
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i + 1); // Column A (index 0)
        }

        // Add a sparkline group of type Line
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // Add a sparkline to column B (index 1) at row 1 (index 0) using the data range A1:A10
        // SparklineCollection.Add(string dataRange, int row, int column)
        sparklineGroup.Sparklines.Add("A1:A10", 0, 1);

        // Save the workbook
        workbook.Save("LineSparkline.xlsx");
    }
}
