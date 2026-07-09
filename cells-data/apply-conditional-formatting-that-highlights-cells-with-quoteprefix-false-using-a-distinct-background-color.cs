using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            // Cell with leading single quote (QuotePrefix = true)
            worksheet.Cells["A1"].PutValue("'12345");
            // Cell without leading single quote (QuotePrefix = false)
            worksheet.Cells["A2"].PutValue("67890");
            // Another cell without leading single quote
            worksheet.Cells["A3"].PutValue("ABC");

            // Add a conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied (A1:A3)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 0
            };
            conditions.AddArea(area);

            // Add an expression‑based condition that evaluates to TRUE when the cell does NOT start with a single quote
            // Formula: =LEFT(A1,1)<>"'"
            // Note: The formula uses relative reference; Aspose.Cells will adjust it for each cell in the range.
            int conditionIndex = conditions.AddCondition(FormatConditionType.Expression, OperatorType.None, "=LEFT(A1,1)<>\"'\"", null);
            FormatCondition condition = conditions[conditionIndex];

            // Set the style to be applied when the condition is met (QuotePrefix = false)
            condition.Style.BackgroundColor = Color.LightGreen;

            // Save the workbook
            workbook.Save("QuotePrefixConditionalFormatting.xlsx");
        }
    }
}