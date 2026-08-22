// Title: Use Aspose.Cells for .NET to apply formula‑based conditional formatting that highlights values above 150 in red and below 30 in light blue in column A
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, fills cells A1:A10 with numeric values, defines a CellArea for that range, adds two expression‑type FormatCondition objects ("=A1>150" and "=A1<30"), sets the first condition’s background to red and the second to light blue, and saves the file. | Generate a complete Aspose.Cells example that demonstrates adding multiple conditional formatting rules based on formulas to a single column, including configuring style colors and persisting the workbook.
// Common Searches: aspnet apply conditional formatting with formula >150 red using Aspose.Cells | c# Aspose.Cells conditional formatting multiple expression conditions column A | how to set background color for cells less than 30 in Aspose.Cells | example of adding formula based conditional formatting to Excel with Aspose.Cells .NET
// Tags: Aspose.Cells expression conditional formatting | C# set cell background based on value | Excel conditional formatting formula thresholds Aspose | Aspose.Cells multiple format conditions | save workbook with conditional formats .NET

using System;
using Aspose.Cells;
using System.Drawing;

// Demonstrates creating a workbook, populating column A with multiples of 30, defining the range A1:A10, adding two expression‑type conditional formatting rules (value >150 with red background, value <30 with light blue background), and saving the result as ThresholdConditionalFormatting.xlsx.
class ConditionalFormattingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample numeric data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue(i * 30); // 0, 30, 60, ... 270
        }

        // Define the cell area to which the conditional formatting will be applied (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];
        fcc.AddArea(area);

        // -------------------------------------------------
        // Condition 1: Highlight cells with value > 150 (red background)
        // -------------------------------------------------
        int conditionIdx1 = fcc.AddCondition(
            FormatConditionType.Expression,
            OperatorType.None,
            "=A1>150",   // Formula evaluates to TRUE when the cell value exceeds 150
            null);       // No second formula needed for Expression type

        FormatCondition condition1 = fcc[conditionIdx1];
        condition1.Style.BackgroundColor = Color.Red;

        // -------------------------------------------------
        // Condition 2: Highlight cells with value < 30 (light blue background)
        // -------------------------------------------------
        int conditionIdx2 = fcc.AddCondition(
            FormatConditionType.Expression,
            OperatorType.None,
            "=A1<30",    // Formula evaluates to TRUE when the cell value is below 30
            null);

        FormatCondition condition2 = fcc[conditionIdx2];
        condition2.Style.BackgroundColor = Color.LightBlue;

        // Save the workbook with the applied conditional formatting
        workbook.Save("ThresholdConditionalFormatting.xlsx");
    }
}
