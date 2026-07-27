// Title: Aspose.Cells C# – Apply a 3‑Arrow Icon Set with Custom Numeric Thresholds
// Description: Demonstrates how to create a workbook, fill A1:A10 with numbers, and add an IconSet conditional formatting rule (Arrows3) that shows green‑up, yellow‑right and red‑down arrows based on minimum, 30 and 70 thresholds, while optionally displaying the cell value.
// Keywords: Aspose.Cells | C# | icon set | arrow icons | conditional formatting | numeric thresholds | IconSetType.Arrows3 | FormatConditionValueType | Excel automation
// Common Searches: Aspose.Cells add arrow icon set C# | icon set conditional formatting example .NET | set numeric thresholds for IconSet Aspose.Cells | show arrows based on cell values using Aspose | reverse order hide value IconSet Aspose.Cells
// Developer Intent: Add an arrow‑based IconSet conditional formatting rule to a specific cell range and define custom numeric cut‑offs.
// Use Cases: Visualize KPI trends with up/down arrows next to each metric. | Highlight sales performance bands: low (red), medium (yellow), high (green). | Create a compact dashboard where values and directional icons are shown together. | Generate automated reports that instantly convey growth or decline.
// AI Prompts: Write C# code with Aspose.Cells to apply a three‑arrow IconSet to cells A1:A20 using thresholds 20 and 80. | Explain how to reverse the arrow order and hide the numeric value in an IconSet rule. | Show how to replace fixed thresholds with percentile‑based thresholds for an IconSet in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsIconSetExample
{
    // Demonstrates how to create a workbook, fill A1:A10 with numbers, and add an IconSet conditional formatting rule (Arrows3) that shows green‑up, yellow‑right and red‑down arrows based on minimum, 30 and 70 thresholds, while optionally displaying the cell value.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (rows 1‑10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 15); // 0,15,30,...,135
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

            // Add an IconSet condition
            int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcs[conditionIdx];

            // Configure the icon set to use three arrows
            condition.IconSet.Type = IconSetType.Arrows3;
            condition.IconSet.ShowValue = true;   // Show the cell value next to the icon
            condition.IconSet.Reverse = false;    // Keep default order (green up, yellow right, red down)

            // Set the three threshold values (Min, Mid, Max)
            // First CFVO – minimum (automatically uses the lowest value in the range)
            condition.IconSet.Cfvos[0].Type = FormatConditionValueType.Min;
            condition.IconSet.Cfvos[0].Value = null;
            condition.IconSet.Cfvos[0].IsGTE = true;

            // Second CFVO – middle threshold (e.g., 30)
            condition.IconSet.Cfvos[1].Type = FormatConditionValueType.Number;
            condition.IconSet.Cfvos[1].Value = 30;
            condition.IconSet.Cfvos[1].IsGTE = true;

            // Third CFVO – maximum threshold (e.g., 70)
            condition.IconSet.Cfvos[2].Type = FormatConditionValueType.Number;
            condition.IconSet.Cfvos[2].Value = 70;
            condition.IconSet.Cfvos[2].IsGTE = true;

            // Save the workbook
            workbook.Save("IconSetArrowsConditionalFormatting.xlsx");
        }
    }
}
