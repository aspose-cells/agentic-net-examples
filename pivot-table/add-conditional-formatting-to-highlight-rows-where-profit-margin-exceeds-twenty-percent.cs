// Title: Apply conditional formatting in Aspose.Cells for .NET to highlight rows where profit margin exceeds 20 percent
// Common Searches: Aspose.Cells C# conditional formatting profit margin greater than 20 percent | How to highlight Excel rows when column C divided by column B is over 0.2 using Aspose.Cells | Set yellow fill for rows based on expression =$C2/$B2>0.2 in Aspose.Cells .NET | Add conditional formatting to range A2:D100 with Aspose.Cells API | Create expression‑based conditional rule in a workbook using Aspose.Cells
// Tags: conditional formatting expression Aspose.Cells | yellow fill style Aspose.Cells | apply conditional rule to cell area | profit margin highlight rows C# | formatconditioncollection addcondition usage

using System;
using System.Drawing;
using Aspose.Cells;
using Range = Aspose.Cells.Range;

// The example creates a workbook, defines the range A2:D100, adds an expression‑based conditional formatting rule ($C2/$B2>0.2), sets a solid yellow background style, applies the rule to the specified range, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range (rows 2‑100, columns A‑D). Indices are zero‑based.
            int firstRow = 1;   // row 2
            int lastRow = 99;   // row 100
            int firstCol = 0;   // column A
            int lastCol = 3;    // column D

            // Build a CellArea representing the range
            CellArea dataArea = new CellArea
            {
                StartRow = firstRow,
                StartColumn = firstCol,
                EndRow = lastRow,
                EndColumn = lastCol
            };

            // Add a new conditional formatting rule collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];

            // Expression: profit margin (C / B) > 20%
            int conditionIndex = conditions.AddCondition(
                FormatConditionType.Expression,
                OperatorType.None,
                "=$C2/$B2>0.2",
                null);
            FormatCondition condition = conditions[conditionIndex];

            // Define the formatting (yellow fill) to apply when the condition is true
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;
            condition.Style = highlightStyle;

            // Apply the conditional formatting to the defined range
            sheet.ConditionalFormattings[cfIndex].AddArea(dataArea);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
