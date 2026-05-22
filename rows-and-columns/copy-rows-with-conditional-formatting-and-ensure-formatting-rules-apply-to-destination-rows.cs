using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCopyRowsWithConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            // Create source workbook and get its first worksheet
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Populate some data in rows 0-4
            for (int row = 0; row < 5; row++)
            {
                sourceSheet.Cells[row, 0].PutValue(row * 10);          // Column A
                sourceSheet.Cells[row, 1].PutValue("Item " + row);    // Column B
            }

            // Add a conditional formatting rule: highlight cells in column A > 20
            int cfIndex = sourceSheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sourceSheet.ConditionalFormattings[cfIndex];
            fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "20", null);
            Style cfStyle = sourceWorkbook.CreateStyle();
            cfStyle.ForegroundColor = Color.Yellow;
            cfStyle.Pattern = BackgroundType.Solid;
            fcc[0].Style = cfStyle;

            // Apply the conditional formatting to the range A1:A5
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Create destination workbook (could be the same workbook with a new sheet)
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Prepare copy options to extend formatting to adjacent range if needed
            CopyOptions copyOptions = new CopyOptions
            {
                ExtendToAdjacentRange = true
            };

            // Copy rows 0-4 from source to destination starting at row 5
            destSheet.Cells.CopyRows(sourceSheet.Cells, 0, 5, 5, copyOptions);

            // Ensure conditional formatting rules are also copied to the destination sheet
            destSheet.ConditionalFormattings.Copy(sourceSheet.ConditionalFormattings);

            // Save the result workbook
            destWorkbook.Save("RowsWithConditionalFormatting.xlsx");
        }
    }
}