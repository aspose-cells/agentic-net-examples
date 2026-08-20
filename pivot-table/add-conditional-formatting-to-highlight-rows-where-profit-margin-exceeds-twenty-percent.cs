// Title: Aspose.Cells C# – Highlight rows where profit margin exceeds 20% with conditional formatting
// Description: Creates a workbook, adds sample revenue data, defines a conditional‑formatting range (A‑D), applies an expression "=D1>0.2" to color the entire row light‑yellow when the profit‑margin column (D) is greater than 20%, and saves the file as ProfitMarginHighlight.xlsx.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | profit margin | highlight rows | Excel expression formula | entire row style | background color | financial reporting
// Common Searches: Aspose.Cells highlight rows profit margin >20% | C# conditional formatting entire row Excel | Aspose.Cells set background color with formula | How to apply expression condition in Aspose.Cells | Conditional formatting range A-D Aspose.Cells .NET
// Developer Intent: Add a conditional‑formatting rule that automatically colors any worksheet row whose profit‑margin value in column D is above 20%.
// Use Cases: Visually flag high‑margin products in sales dashboards. | Color‑code rows that meet profitability thresholds in financial statements. | Generate invoices that automatically highlight lucrative transactions.
// AI Prompts: Write C# code using Aspose.Cells to apply conditional formatting that colors rows where column D (profit margin) > 0.2. | Show how to change the margin threshold or target a different column in the conditional‑formatting rule. | Demonstrate adding multiple rules (e.g., low margin red, high margin green) to the same worksheet with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, adds sample revenue data, defines a conditional‑formatting range (A‑D), applies an expression "=D1>0.2" to color the entire row light‑yellow when the profit‑margin column (D) is greater than 20%, and saves the file as ProfitMarginHighlight.xlsx.
class HighlightProfitMargin
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional)
        // Column A: Item, B: Revenue, C: Cost, D: Profit Margin
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Revenue");
        sheet.Cells["C1"].PutValue("Cost");
        sheet.Cells["D1"].PutValue("ProfitMargin");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["C2"].PutValue(700);
        sheet.Cells["D2"].PutValue(0.30); // 30%
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["C3"].PutValue(600);
        sheet.Cells["D3"].PutValue(0.15); // 15%
        // Add more rows as needed...

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range to which the formatting will be applied (e.g., rows 1-100, columns A-D)
        CellArea area = new CellArea
        {
            StartRow = 1,      // Row 2 in Excel (zero‑based index)
            EndRow = 100,      // Adjust as needed
            StartColumn = 0,   // Column A
            EndColumn = 3      // Column D
        };
        fcs.AddArea(area);

        // Add an expression‑type condition
        int condIdx = fcs.AddCondition(FormatConditionType.Expression);
        FormatCondition fc = fcs[condIdx];

        // Formula evaluates the profit margin in column D of the current row
        // The relative reference (D1) will adjust for each row in the range
        fc.Formula1 = "=D1>0.2";

        // Set the style to highlight the entire row (e.g., light yellow background)
        fc.Style.BackgroundColor = Color.LightYellow;

        // Save the workbook
        workbook.Save("ProfitMarginHighlight.xlsx");
    }
}
