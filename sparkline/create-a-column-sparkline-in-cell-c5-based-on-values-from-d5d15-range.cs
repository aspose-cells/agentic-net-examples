// Title: Create a column sparkline in cell C5 from the D5:D15 range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a column‑type sparkline to cell C5 referencing the data in D5:D15 and saves the workbook as an .xlsx file. | Demonstrate how to configure a SparklineGroup of type Column, place a sparkline at C5, and export the worksheet using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to insert a column sparkline at a specific cell | example code for creating a sparkline from D5 to D15 in Aspose.Cells | C# add column sparkline to worksheet using Aspose.Cells library | save workbook with sparkline in Aspose.Cells .NET | populate data and generate column sparkline in Aspose.Cells C# tutorial
// Tags: Aspose.Cells column sparkline group | C# add sparkline to cell | sparkline from range D5:D15 | export workbook with sparkline .xlsx | Aspose.Cells SparklineType.Column usage

using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells D5:D15 with sample values, adds a column‑type SparklineGroup, places a sparkline in cell C5 that references the D5:D15 range, and saves the file as ColumnSparkline.xlsx.
class SparklineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the range D5:D15 (optional, for demonstration)
        for (int i = 0; i < 11; i++)
        {
            sheet.Cells[4 + i, 3].PutValue(i + 1); // Row 4+i (5th to 15th), Column 3 (D)
        }

        // Add a sparkline group of type Column
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group:
        // Data range: D5:D15
        // Location: row 4 (C5), column 2 (C)
        group.Sparklines.Add("D5:D15", 4, 2);

        // Save the workbook
        workbook.Save("ColumnSparkline.xlsx");
    }
}
