using System;
using Aspose.Cells;
using Aspose.Cells;

namespace AsposeCellsIconSetExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (rows 0‑9)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
            }

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the icon set will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add an IconSet conditional formatting rule
            int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcs[conditionIdx];

            // Configure the icon set to use the 3‑arrow set
            condition.IconSet.Type = IconSetType.Arrows3;
            condition.IconSet.ShowValue = true;   // Show the cell value alongside the icon
            condition.IconSet.Reverse = false;    // Keep default icon order

            // Set numeric thresholds for the three icons
            // First threshold (lowest icon)
            condition.IconSet.Cfvos[0].Type = FormatConditionValueType.Number;
            condition.IconSet.Cfvos[0].Value = 0;
            condition.IconSet.Cfvos[0].IsGTE = true;

            // Second threshold (middle icon)
            condition.IconSet.Cfvos[1].Type = FormatConditionValueType.Number;
            condition.IconSet.Cfvos[1].Value = 50;
            condition.IconSet.Cfvos[1].IsGTE = true;

            // Third threshold (highest icon)
            condition.IconSet.Cfvos[2].Type = FormatConditionValueType.Number;
            condition.IconSet.Cfvos[2].Value = 100;
            condition.IconSet.Cfvos[2].IsGTE = true;

            // Save the workbook
            workbook.Save("IconSetArrowsConditionalFormatting.xlsx");
        }
    }
}