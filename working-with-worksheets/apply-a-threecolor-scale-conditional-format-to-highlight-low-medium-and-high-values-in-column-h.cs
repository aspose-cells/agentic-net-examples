using System;
using System.Drawing;
using Aspose.Cells;

class ThreeColorScaleExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Populate column H with sample numeric data
        for (int row = 0; row < 100; row++)
        {
            worksheet.Cells[row, 7].PutValue(row); // Column H is index 7 (zero‑based)
        }

        // Add a new conditional formatting collection
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

        // Define the range for column H (rows 0‑99)
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = 99,
            StartColumn = 7,
            EndColumn = 7
        };
        cfCollection.AddArea(range);

        // Add a ColorScale condition
        int conditionIndex = cfCollection.AddCondition(FormatConditionType.ColorScale);
        FormatCondition condition = cfCollection[conditionIndex];

        // Configure the three‑color scale (green → yellow → red)
        ColorScale scale = condition.ColorScale;
        scale.Is3ColorScale = true;
        scale.MinColor = Color.Green;
        scale.MidColor = Color.Yellow;
        scale.MaxColor = Color.Red;

        // Set the value types for min, mid, and max
        scale.MinCfvo.Type = FormatConditionValueType.Min;          // Minimum value in the range
        scale.MidCfvo.Type = FormatConditionValueType.Percentile; // 50th percentile (median)
        scale.MidCfvo.Value = 50;
        scale.MaxCfvo.Type = FormatConditionValueType.Max;          // Maximum value in the range

        // Save the workbook
        workbook.Save("ThreeColorScaleColumnH.xlsx");
    }
}