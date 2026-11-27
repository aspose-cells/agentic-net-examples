// Title: Apply conditional formatting in C# with Aspose.Cells to highlight column F cells that exceed the column’s average value
// Common Searches: C# Aspose.Cells conditional formatting highlight cells greater than column average | How to set an expression rule for average value in Aspose.Cells .NET | Apply style to Excel column F based on AVERAGE formula using Aspose.Cells | Aspose.Cells .NET add conditional formatting for a specific column range | Create conditional formatting rule with expression in Aspose.Cells C# example
// Tags: Aspose.Cells conditional formatting expression | highlight cells above column average in Excel | C# set conditional formatting range F1:F10 | apply light salmon fill style Aspose.Cells | average formula AVERAGE($F:$F) Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a workbook, populates cells F1‑F10, adds a conditional formatting collection for the range F1:F10, defines an expression condition `F1>AVERAGE($F:$F)` (adjusted per row) to compare each cell with the column’s average, applies a LightSalmon solid fill style when true, and saves the file as ConditionalFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // (Optional) Populate column F with sample data for demonstration
            for (int i = 0; i < 10; i++)
            {
                // Column F has index 5 (0‑based)
                sheet.Cells[i, 5].PutValue(i * 10 + 5);
            }

            // Define the range to which the conditional formatting will be applied (F1:F10)
            int firstRow = 0;          // Row 1 (0‑based)
            int lastRow = 9;           // Row 10
            int columnIndex = 5;       // Column F

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Specify the area (range) for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = columnIndex,
                EndRow = lastRow,
                EndColumn = columnIndex
            };
            fcc.AddArea(area);

            // Add a condition: cell value > AVERAGE of the entire column F
            // Use an expression formula; the row number will adjust for each cell in the range
            string formula = $"F{firstRow + 1}>AVERAGE($F:$F)";
            int conditionIndex = fcc.AddCondition(FormatConditionType.Expression,
                                                   OperatorType.None,
                                                   null,
                                                   formula);

            // Define the style to apply when the condition is true (e.g., light red fill)
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.LightSalmon;
            highlightStyle.Pattern = BackgroundType.Solid;
            fcc[conditionIndex].Style = highlightStyle;

            // Save the workbook (lifecycle rule)
            workbook.Save("ConditionalFormatting.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
