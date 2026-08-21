// Title: C# – Aspose.Cells: Apply a Three‑Color Scale Conditional Formatting to Column H
// Description: Creates a new workbook, defines a CellArea for rows 1‑100 of column H, adds a three‑color scale (LightBlue = low, Yellow = median, OrangeRed = high) using Aspose.Cells for .NET, and saves the file as ThreeColorScale_ColumnH.xlsx.
// Keywords: Aspose.Cells C# three color scale | conditional formatting column H | Excel color scale .NET | FormatCondition ColorScale Aspose | C# Excel gradient formatting example
// Common Searches: Aspose.Cells three color scale column H C# | how to add percentile based color scale with Aspose.Cells | C# code for conditional formatting Excel column H | apply gradient conditional format to a range using Aspose.Cells
// Developer Intent: Generate a three‑color scale conditional formatting rule for column H in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Visually differentiate low, medium, and high performance metrics in a report column. | Show sales or KPI trends across the first 100 rows with a gradient from LightBlue to OrangeRed. | Automatically color‑code data‑driven dashboards where the middle color reflects the 50th percentile.
// AI Prompts: Write C# code with Aspose.Cells to apply a three‑color scale to column H (rows 1‑100) using LightBlue, Yellow, and OrangeRed. | Explain how to modify the middle percentile value in an Aspose.Cells three‑color scale rule. | Provide an example that applies a three‑color scale to a dynamic range in column H using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormatting
{
    // Creates a new workbook, defines a CellArea for rows 1‑100 of column H, adds a three‑color scale (LightBlue = low, Yellow = median, OrangeRed = high) using Aspose.Cells for .NET, and saves the file as ThreeColorScale_ColumnH.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range for column H (zero‑based column index 7)
            // Here we apply the format to rows 0 through 99 (adjust as needed)
            CellArea range = new CellArea
            {
                StartRow = 0,
                EndRow = 99,
                StartColumn = 7,
                EndColumn = 7
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

            // Associate the defined range with the conditional formatting
            cfCollection.AddArea(range);

            // Add a ColorScale condition
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.ColorScale);
            FormatCondition condition = cfCollection[conditionIndex];

            // Configure the three‑color scale
            ColorScale colorScale = condition.ColorScale;
            colorScale.Is3ColorScale = true;               // Enable three‑color scale
            colorScale.MinColor = Color.LightBlue;         // Low values color
            colorScale.MidColor = Color.Yellow;            // Mid values color
            colorScale.MaxColor = Color.OrangeRed;         // High values color

            // Set the value types for min, mid, and max
            colorScale.MinCfvo.Type = FormatConditionValueType.Min;          // Minimum of the range
            colorScale.MidCfvo.Type = FormatConditionValueType.Percentile;   // 50th percentile (median)
            colorScale.MidCfvo.Value = 50;
            colorScale.MaxCfvo.Type = FormatConditionValueType.Max;          // Maximum of the range

            // Save the workbook
            workbook.Save("ThreeColorScale_ColumnH.xlsx");
        }
    }
}
