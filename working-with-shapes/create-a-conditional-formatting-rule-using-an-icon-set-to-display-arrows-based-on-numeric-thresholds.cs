// Title: Apply a 3‑Arrow Icon Set Conditional Formatting Rule with Aspose.Cells for C#
// Description: Demonstrates how to create a workbook, populate cells A1‑A4, add a conditional‑formatting collection, and assign a 3‑arrow IconSet (Arrows3) with numeric thresholds (0, 50, 100). The example shows the cell value beside each arrow and saves the result as an XLSX file.
// Keywords: Aspose.Cells C# | icon set conditional formatting | 3‑arrow icon set | Excel conditional formatting Aspose | numeric thresholds icon set | Arrows3 Aspose.Cells | C# Excel automation | Aspose.Cells tutorial | conditional formatting arrows .NET | GitHub Aspose.Cells examples
// Common Searches: Aspose.Cells add 3‑arrow icon set | C# conditional formatting icon set thresholds | How to use IconSetType.Arrows3 in Aspose.Cells | Display arrows based on numeric values Excel C# | Aspose.Cells conditional formatting range example
// Developer Intent: Create a conditional‑formatting rule that visualizes numeric values with a three‑arrow icon set in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Show sales growth direction with up, right, and down arrows. | Visualize project milestone status in a dashboard worksheet. | Map risk scores to low, medium, and high arrows for quick assessment.
// AI Prompts: Generate C# code to apply a 4‑traffic‑light icon set with custom numeric thresholds using Aspose.Cells. | Explain how to reverse the order of an icon set in Aspose.Cells conditional formatting. | Provide sample code that hides the cell value and displays only the icon in an Aspose.Cells icon set.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsIconSetExample
{
    // Demonstrates how to create a workbook, populate cells A1‑A4, add a conditional‑formatting collection, and assign a 3‑arrow IconSet (Arrows3) with numeric thresholds (0, 50, 100). The example shows the cell value beside each arrow and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(45);
            sheet.Cells["A3"].PutValue(75);
            sheet.Cells["A4"].PutValue(110);

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the icon set will be applied (A1:A4)
            CellArea area = CellArea.CreateCellArea("A1", "A4");
            fcs.AddArea(area);

            // Add an IconSet condition
            int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcs[conditionIdx];

            // Configure the icon set to use the 3‑arrow set
            condition.IconSet.Type = IconSetType.Arrows3;
            condition.IconSet.ShowValue = true;   // Show the cell value alongside the icon
            condition.IconSet.Reverse = false;    // Keep default icon order

            // Define the three threshold values for the arrows
            // First threshold (lowest) – values <= 0 show the lowest arrow
            condition.IconSet.Cfvos[0].Type = FormatConditionValueType.Number;
            condition.IconSet.Cfvos[0].Value = "0";
            condition.IconSet.Cfvos[0].IsGTE = true;

            // Second threshold – values > 0 and <= 50 show the middle arrow
            condition.IconSet.Cfvos[1].Type = FormatConditionValueType.Number;
            condition.IconSet.Cfvos[1].Value = "50";
            condition.IconSet.Cfvos[1].IsGTE = true;

            // Third threshold – values > 50 show the highest arrow
            condition.IconSet.Cfvos[2].Type = FormatConditionValueType.Number;
            condition.IconSet.Cfvos[2].Value = "100";
            condition.IconSet.Cfvos[2].IsGTE = true;

            // Save the workbook to an XLSX file
            workbook.Save("IconSetArrowsExample.xlsx");
        }
    }
}
