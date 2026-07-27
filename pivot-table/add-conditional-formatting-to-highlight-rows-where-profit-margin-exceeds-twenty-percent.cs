using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: columns A (Product), B (Revenue), C (Cost), D (Profit Margin)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["C1"].PutValue("Cost");
            sheet.Cells["D1"].PutValue("ProfitMargin");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["C2"].PutValue(800);
            sheet.Cells["D2"].PutValue(0.33); // 33%

            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(900);
            sheet.Cells["C3"].PutValue(750);
            sheet.Cells["D3"].PutValue(0.17); // 17%

            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(1500);
            sheet.Cells["C4"].PutValue(1000);
            sheet.Cells["D4"].PutValue(0.25); // 25%

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will be applied (rows 2-4, columns A-D)
            CellArea area = new CellArea
            {
                StartRow = 1,   // zero‑based index (row 2 in Excel)
                EndRow = 3,     // row 4
                StartColumn = 0,
                EndColumn = 3   // column D
            };
            fcs.AddArea(area);

            // Add an expression‑type condition
            int conditionIndex = fcs.AddCondition(FormatConditionType.Expression);
            FormatCondition fc = fcs[conditionIndex];

            // Formula checks if the profit margin in column D of the current row exceeds 20%
            fc.Formula1 = "=$D2>0.2";

            // Set the style to highlight the entire row
            fc.Style.BackgroundColor = Color.LightGreen;

            // Save the workbook
            workbook.Save("ProfitMarginConditionalFormatting.xlsx");
        }
    }
}