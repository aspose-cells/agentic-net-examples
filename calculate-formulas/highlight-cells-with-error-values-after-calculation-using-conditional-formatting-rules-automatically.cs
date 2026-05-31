using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsErrorHighlightDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Insert formulas that will generate errors
            cells["A1"].Formula = "=1/0";                     // Division by zero error
            cells["B1"].Formula = "=SUM(A1,5)";               // Propagates error from A1
            cells["C1"].Formula = "=UNKNOWNFUNC()";           // Unsupported function error

            // Calculate all formulas so that error values are materialized
            workbook.CalculateFormula();

            // Determine the used range to apply conditional formatting
            int lastRow = cells.MaxDataRow;
            int lastColumn = cells.MaxDataColumn;
            CellArea usedRange = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = lastRow,
                EndColumn = lastColumn
            };

            // Add a new conditional formatting collection
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

            // Apply the range to the conditional formatting
            cfCollection.AddArea(usedRange);

            // Add a condition that highlights cells containing errors
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.ContainsErrors);
            FormatCondition condition = cfCollection[conditionIndex];

            // Set the highlight style (background color)
            condition.Style.BackgroundColor = Color.LightPink;

            // Save the workbook
            workbook.Save("ErrorHighlighted.xlsx");
        }
    }
}