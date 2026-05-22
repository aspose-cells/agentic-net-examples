using System;
using System.Drawing;
using Aspose.Cells;

namespace ProgressBarConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample progress values (0-100)
            double[] progressValues = { 10, 35, 55, 75, 90 };
            for (int i = 0; i < progressValues.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(progressValues[i]);
            }

            // Define the range that will hold the progress bars
            CellArea range = new CellArea
            {
                StartRow = 0,
                EndRow = progressValues.Length - 1,
                StartColumn = 0,
                EndColumn = 0
            };

            // Add an empty conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
            cfCollection.AddArea(range);

            // ---------- Low range (0 - 30) : Red ----------
            int lowCondIdx = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition lowCond = cfCollection[lowCondIdx];
            lowCond.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
            lowCond.DataBar.MinCfvo.Value = 0;
            lowCond.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
            lowCond.DataBar.MaxCfvo.Value = 30;
            lowCond.DataBar.Color = Color.Red;
            lowCond.DataBar.ShowValue = true; // show the numeric value

            // ---------- Mid range (31 - 70) : Orange ----------
            int midCondIdx = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition midCond = cfCollection[midCondIdx];
            midCond.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
            midCond.DataBar.MinCfvo.Value = 31;
            midCond.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
            midCond.DataBar.MaxCfvo.Value = 70;
            midCond.DataBar.Color = Color.Orange;
            midCond.DataBar.ShowValue = true;

            // ---------- High range (71 - 100) : Green ----------
            int highCondIdx = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition highCond = cfCollection[highCondIdx];
            highCond.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
            highCond.DataBar.MinCfvo.Value = 71;
            highCond.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
            highCond.DataBar.MaxCfvo.Value = 100;
            highCond.DataBar.Color = Color.Green;
            highCond.DataBar.ShowValue = true;

            // Save the workbook
            workbook.Save("ProgressBarConditionalFormatting.xlsx", SaveFormat.Xlsx);
        }
    }
}