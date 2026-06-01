using System;
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

            // Populate sample numeric values in the target range C3:C12
            for (int row = 2; row <= 11; row++) // rows are zero‑based
            {
                sheet.Cells[row, 2].PutValue(row - 1); // C column index = 2
            }

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the cell area C3:C12 (rows 2‑11, column 2)
            CellArea area = new CellArea
            {
                StartRow = 2,
                EndRow = 11,
                StartColumn = 2,
                EndColumn = 2
            };
            fcs.AddArea(area);

            // Add an IconSet condition
            int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcs[conditionIdx];

            // Choose an icon set type (any built‑in type)
            condition.IconSet.Type = IconSetType.TrafficLights31;

            // Ensure the cell values are displayed (text labels) alongside the icons
            condition.IconSet.ShowValue = true;

            // Replace each default icon with a "None" type so only the text label remains visible
            foreach (ConditionalFormattingIcon icon in condition.IconSet.CfIcons)
            {
                icon.Type = IconSetType.None; // No icon will be rendered
                icon.Index = 0;               // Index is irrelevant for None type
            }

            // Save the workbook
            workbook.Save("IconSetWithTextLabels.xlsx");
        }
    }
}