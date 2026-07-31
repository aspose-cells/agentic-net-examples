// Title: Add a Three‑Icon Traffic‑Light Conditional Format to Column M with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, selects rows 0‑99 of column M, adds an IconSet conditional format of type TrafficLights31, defines low and medium thresholds at 33 % and 67 % using percentage values, and saves the result as ThreeIconSet.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | C# | conditional formatting | icon set | traffic lights | three icon set | column M | percentage thresholds | Excel workbook | FormatCondition | IconSetType.TrafficLights31
// Common Searches: Aspose.Cells three icon set column M C# | add traffic‑light conditional formatting with Aspose.Cells | set percentage thresholds for IconSet in .NET | C# example conditional formatting icon set Excel | apply IconSetType.TrafficLights31 to a range using Aspose.Cells
// Developer Intent: Apply a three‑icon traffic‑light conditional format to column M so that low, medium, and high values are shown with red, yellow, and green icons.
// Use Cases: Highlight performance scores in a KPI report with red‑yellow‑green icons. | Provide instant visual cues for sales or rating columns in a dashboard spreadsheet. | Automatically format newly generated data rows for quick trend analysis.
// AI Prompts: Generate C# code that uses Aspose.Cells to apply a four‑icon set with custom absolute thresholds to a specified range. | Show how to change the example to use the 'Arrows3' icon set and numeric value thresholds instead of percentages. | Explain how to loop through multiple columns and apply the same three‑icon traffic‑light conditional format with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells;

// C# example that creates a workbook, selects rows 0‑99 of column M, adds an IconSet conditional format of type TrafficLights31, defines low and medium thresholds at 33 % and 67 % using percentage values, and saves the result as ThreeIconSet.xlsx using Aspose.Cells.
class ThreeIconSetConditionalFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range: entire column M (zero‑based column index 12), rows 0‑99
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 99,
            StartColumn = 12,
            EndColumn = 12
        };
        fcs.AddArea(area);

        // Add an IconSet condition
        int conditionIndex = fcs.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcs[conditionIndex];

        // Choose a three‑icon set (Traffic Lights: red, yellow, green)
        condition.IconSet.Type = IconSetType.TrafficLights31;

        // Configure the three thresholds (low, medium, high) using percentages
        // First threshold – low (0 % – 33 %)
        condition.IconSet.Cfvos[0].Type = FormatConditionValueType.Percent;
        condition.IconSet.Cfvos[0].Value = 33;
        condition.IconSet.Cfvos[0].IsGTE = true; // greater‑than‑or‑equal (default)

        // Second threshold – medium (33 % – 67 %)
        condition.IconSet.Cfvos[1].Type = FormatConditionValueType.Percent;
        condition.IconSet.Cfvos[1].Value = 67;
        condition.IconSet.Cfvos[1].IsGTE = true;

        // The third threshold (high) is implicit as the maximum value

        // Save the workbook
        workbook.Save("ThreeIconSet.xlsx", SaveFormat.Xlsx);
    }
}
