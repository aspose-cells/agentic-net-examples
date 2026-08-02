// Title: C# – Apply a Three‑Color Scale Conditional Formatting to Column H with Aspose.Cells for .NET
// Description: Creates a new Workbook, optionally fills column H with numeric data, defines a CellArea covering the used rows, adds a ColorScale conditional format, configures a red‑yellow‑green three‑color gradient (Min‑Percentile‑Max), and saves the file as ThreeColorScale.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | three color scale | conditional formatting | Excel | column H | ColorScale | Min Percentile Max | red yellow green | sample code | Workbook | Worksheet | CellArea
// Common Searches: Aspose.Cells three color scale column H C# | how to add conditional formatting with three colors in Excel using Aspose.Cells | C# code for red yellow green color scale in Excel | set percentile based color scale Aspose.Cells .NET | apply three‑color conditional format to a column with Aspose.Cells
// Developer Intent: Add a three‑color scale conditional format to column H so low, medium, and high numeric values are highlighted with red, yellow, and green gradients.
// Use Cases: Highlight sales totals in column H with a red‑yellow‑green gradient to spot under‑performing and top‑performing items. | Display employee performance scores in generated reports by applying a three‑color scale to the score column. | Color‑code risk scores in column H of an exported workbook for quick visual risk assessment.
// AI Prompts: Show how to change the three‑color scale to blue, white, and red while keeping the same thresholds. | Generate C# code that applies a two‑color scale conditional format to columns A‑C using Aspose.Cells. | Explain how to use custom numeric thresholds instead of Min/Percentile/Max for a three‑color scale in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new Workbook, optionally fills column H with numeric data, defines a CellArea covering the used rows, adds a ColorScale conditional format, configures a red‑yellow‑green three‑color gradient (Min‑Percentile‑Max), and saves the file as ThreeColorScale.xlsx using Aspose.Cells for .NET.
class ThreeColorScaleConditionalFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Populate column H with sample numeric data
        for (int row = 0; row < 20; row++)
        {
            worksheet.Cells[row, 7].PutValue(row + 1); // Column H has index 7 (0‑based)
        }

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

        // Define the range for column H (from row 0 to the last used row)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = worksheet.Cells.MaxDataRow,
            StartColumn = 7,
            EndColumn = 7
        };
        fcs.AddArea(area);

        // Add a ColorScale condition
        int conditionIndex = fcs.AddCondition(FormatConditionType.ColorScale);
        FormatCondition fc = fcs[conditionIndex];

        // Configure the three‑color scale (low = Red, middle = Yellow, high = Green)
        fc.ColorScale.Is3ColorScale = true;
        fc.ColorScale.MinColor = Color.Red;
        fc.ColorScale.MidColor = Color.Yellow;
        fc.ColorScale.MaxColor = Color.Green;

        // Set the value types for min, mid, and max
        fc.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;          // Minimum value in the range
        fc.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile; // 50th percentile (median)
        fc.ColorScale.MidCfvo.Value = 50;
        fc.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;          // Maximum value in the range

        // Save the workbook
        workbook.Save("ThreeColorScale.xlsx");
    }
}
