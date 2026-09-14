// Title: Apply conditional formatting in Aspose.Cells for .NET to highlight rows where profit margin exceeds 20%
// AI Prompts: Generate C# code using Aspose.Cells to add an expression‑based conditional format that colors entire rows light green when the value in column D is greater than 0.20. | Write a C# example that creates a workbook, defines a CellArea covering rows 2‑100 and columns A‑D, applies a conditional formatting rule with formula =$D2>0.20, and saves the file.
// Common Searches: Aspose.Cells C# how to set conditional formatting for an entire row based on a column value | C# Aspose.Cells conditional formatting formula to highlight rows with profit margin over 20 percent | Apply expression condition in Aspose.Cells to color rows where margin column > 0.2
// Tags: conditional formatting expression Aspose.Cells C# | highlight rows based on column value Aspose.Cells | apply style to CellArea Excel .NET | profit margin conditional format Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, populates sample product data with a profit‑margin column, defines a CellArea covering rows 2‑100 and columns A‑D, adds an expression‑based conditional formatting rule (=$D2>0.20) that applies a light‑green style to the entire row, and saves the workbook as ProfitMarginConditionalFormatting.xlsx.
class ProfitMarginConditionalFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Example data: columns A-C are data, column D contains profit margin as a decimal (e.g., 0.25 for 25%)
        // Populate some sample rows (optional, for demonstration)
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Revenue");
        sheet.Cells["C1"].PutValue("Cost");
        sheet.Cells["D1"].PutValue("Margin");
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["C2"].PutValue(700);
        sheet.Cells["D2"].PutValue(0.30); // 30%
        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["C3"].PutValue(600);
        sheet.Cells["D3"].PutValue(0.15); // 15%
        // Add more rows as needed...

        // Define the range to which the conditional formatting will be applied (e.g., rows 2-100, columns A-D)
        int startRow = 1;      // zero‑based index (row 2 in Excel)
        int endRow = 99;       // row 100
        int startColumn = 0;   // column A
        int endColumn = 3;     // column D

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Set the area for the conditional formatting
        CellArea area = new CellArea
        {
            StartRow = startRow,
            EndRow = endRow,
            StartColumn = startColumn,
            EndColumn = endColumn
        };
        fcc.AddArea(area);

        // Add an expression‑type condition (formula based)
        int conditionIdx = fcc.AddCondition(FormatConditionType.Expression);
        FormatCondition condition = fcc[conditionIdx];

        // Formula: highlight the entire row when the margin in column D of that row > 0.20 (20%)
        // Use relative row reference so it works for each row in the range
        condition.Formula1 = "=$D1>0.20";

        // Define the style to apply (e.g., light green background)
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightGreen;
        style.Pattern = BackgroundType.Solid;
        condition.Style = style;

        // Save the workbook
        workbook.Save("ProfitMarginConditionalFormatting.xlsx");
    }
}
