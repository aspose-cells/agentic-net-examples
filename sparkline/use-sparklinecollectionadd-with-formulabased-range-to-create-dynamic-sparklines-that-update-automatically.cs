// Title: Generate a line sparkline in Aspose.Cells .NET that automatically expands using an INDEX‑COUNTA formula
// AI Prompts: Create a workbook, add a line sparkline to cell B1, assign its DataRange to an INDEX‑COUNTA formula that references column A, and save the file. | Append extra rows to column A, trigger formula recalculation, and save a second workbook to show the sparkline updating without manual changes.
// Common Searches: Aspose.Cells set sparkline data range with Excel formula | How to use INDEX and COUNTA for dynamic sparkline range in C# | Automatically update sparkline when adding rows in Aspose.Cells | C# example of formula‑based sparkline that expands with new data | Recalculate formulas after modifying worksheet in Aspose.Cells
// Tags: set sparkline DataRange with INDEX formula Aspose.Cells | dynamic line sparkline using COUNTA in .NET | SparklineGroup.Add line type Aspose.Cells | recalculate workbook formulas after data append | auto‑expanding sparkline range Excel formula

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills column A, adds a line sparkline in B1, and sets its DataRange to an INDEX‑COUNTA formula that grows as rows are added. After saving the initial file, more values are appended, formulas are recalculated, and a second file is saved, demonstrating the sparkline automatically reflecting the expanded data range.
class DynamicSparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate initial data in column A (rows 1‑5)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1); // A1:A5 = 1,2,3,4,5
        }

        // Add a sparkline group of type Line
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline at cell B1 (row 0, column 1) with a temporary static range
        int sparklineIndex = group.Sparklines.Add("A1:A1", 0, 1);
        Sparkline sparkline = group.Sparklines[sparklineIndex];

        // Define a formula‑based dynamic range using INDEX and COUNTA.
        // This range expands automatically as new values are added to column A.
        string dynamicRange = $"{sheet.Name}!A1:INDEX({sheet.Name}!A:A, COUNTA({sheet.Name}!A:A))";

        // Assign the dynamic range to the sparkline
        sparkline.DataRange = dynamicRange;

        // Save the workbook after the initial setup
        workbook.Save("DynamicSparkline_Initial.xlsx");

        // Append more data to column A (rows 6‑10)
        for (int i = 5; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1); // A6:A10 = 6,7,8,9,10
        }

        // Recalculate formulas (if any) to ensure the dynamic range reflects new data
        workbook.CalculateFormula();

        // Save the workbook again – the sparkline in B1 now reflects the expanded range A1:A10
        workbook.Save("DynamicSparkline_Updated.xlsx");
    }
}
