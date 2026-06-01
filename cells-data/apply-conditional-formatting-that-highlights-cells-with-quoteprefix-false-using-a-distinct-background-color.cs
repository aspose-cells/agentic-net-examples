using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Cell A1: value with QuotePrefix = true (starts with a single quote)
            Cell cellA1 = cells["A1"];
            cellA1.PutValue("'12345");
            Style styleA1 = workbook.CreateStyle();
            styleA1.QuotePrefix = true;               // Mark as having a leading quote
            cellA1.SetStyle(styleA1);

            // Cell A2: normal numeric value (QuotePrefix = false)
            cells["A2"].PutValue(12345);

            // Cell A3: text without leading quote (QuotePrefix = false)
            cells["A3"].PutValue("Sample Text");

            // Cell A4: value with QuotePrefix = true (explicitly set via style)
            Cell cellA4 = cells["A4"];
            cellA4.PutValue("67890");
            Style styleA4 = workbook.CreateStyle();
            styleA4.QuotePrefix = true;
            cellA4.SetStyle(styleA4);

            // Define the range to which the conditional formatting will be applied (A1:A4)
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 3,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add an expression‑type condition that evaluates to TRUE when QuotePrefix is FALSE.
            // The formula checks that the first character of the cell is NOT a single quote.
            // LEFT(A1,1)<>\"'\" will be true for cells without a leading quote.
            int conditionIdx = fcs.AddCondition(FormatConditionType.Expression, OperatorType.None,
                                                "LEFT(A1,1)<>\"'\"", null);
            FormatCondition condition = fcs[conditionIdx];

            // Set a distinct background color for cells that meet the condition (QuotePrefix = false)
            condition.Style.BackgroundColor = Color.LightGreen;
            condition.Style.Pattern = BackgroundType.Solid;

            // Save the workbook
            workbook.Save("QuotePrefixConditionalFormatting.xlsx");
        }
    }
}