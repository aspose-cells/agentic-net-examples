// Title: Apply a Light‑Blue Data Bar Conditional Format to Column G with Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, writes incremental numbers to cells G1‑G10, defines a CellArea for column G, adds a DataBar conditional format with automatic minimum and maximum values, sets the bar color to LightBlue, shows the numeric value, and saves the result as ColumnGDataBar.xlsx.
// Keywords: Aspose.Cells | C# | DataBar | ConditionalFormatting | Column G | Excel workbook | AutomaticMin | AutomaticMax | LightBlue bar | SaveFormat.Xlsx | GitHub sample
// Common Searches: Aspose.Cells add data bar to a column C# | conditional formatting data bar column G .NET | automatic min max data bar Aspose.Cells example | set data bar color LightBlue Aspose.Cells | show values with data bar conditional format
// Developer Intent: Add a data‑bar conditional format to column G to visualize numeric trends in an Excel file using Aspose.Cells for .NET.
// Use Cases: Show sales figures in column G with a light‑blue bar that scales from the lowest to highest value. | Display progress percentages in column G while keeping the numeric value visible beside the bar. | Create a quick KPI dashboard where column G values are highlighted with data bars for instant trend analysis.
// AI Prompts: Generate C# code that applies a red DataBar with a custom range (0‑100) to column H using Aspose.Cells. | Modify the sample to hide the numeric values while keeping the LightBlue DataBar visible in column G. | Provide an example that combines a DataBar and a three‑color scale conditional format on the same column with Aspose.Cells.

using System;
using Aspose.Cells;
using System.Drawing;

// This example creates a new workbook, writes incremental numbers to cells G1‑G10, defines a CellArea for column G, adds a DataBar conditional format with automatic minimum and maximum values, sets the bar color to LightBlue, shows the numeric value, and saves the result as ColumnGDataBar.xlsx.
class DataBarColumnG
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample numeric data in column G (index 6)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 6].PutValue(i * 10 + 5);
        }

        // Add a data bar conditional formatting to column G
        int fmtIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[fmtIndex];

        // Define the range for column G (rows 0‑9)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 6,
            EndColumn = 6
        };
        fcs.AddArea(area);

        // Add the DataBar condition
        int condIndex = fcs.AddCondition(FormatConditionType.DataBar);
        FormatCondition condition = fcs[condIndex];

        // Configure the DataBar properties
        DataBar dataBar = condition.DataBar;
        dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
        dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
        dataBar.Color = Color.LightBlue;
        dataBar.ShowValue = true;

        // Save the workbook
        workbook.Save("ColumnGDataBar.xlsx", SaveFormat.Xlsx);
    }
}
