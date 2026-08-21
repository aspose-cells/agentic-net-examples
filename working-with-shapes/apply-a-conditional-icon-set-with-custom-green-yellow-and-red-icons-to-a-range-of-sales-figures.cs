// Title: Apply a Traffic‑Lights (green‑yellow‑red) Icon Set with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills cells A1:A10 with incremental sales values, defines a conditional‑formatting rule, adds the built‑in TrafficLights31 icon set, optionally displays cell values, keeps the default green‑yellow‑red order, customizes each icon's type and index, and saves the result as SalesIconSet.xlsx.
// Keywords: Aspose.Cells | C# conditional formatting | icon set | TrafficLights31 | Excel icon set programmatically | conditional icons | sales dashboard | Excel automation | Aspose.Cells API | conditional formatting icons
// Common Searches: Aspose.Cells add traffic lights icon set C# | conditional icon set Excel using Aspose.Cells .NET | customize icon indices Aspose.Cells | apply three‑color icon set to range Aspose.Cells | show values with icon set Aspose.Cells
// Developer Intent: Add a three‑color traffic‑lights icon set to cells A1:A10 to visually rank sales figures.
// Use Cases: Visualize low, medium, and high sales in a report with green, yellow, and red icons. | Build KPI dashboards where icon sets flag threshold breaches. | Generate automated Excel reports that include conditional icon formatting for quick visual analysis.
// AI Prompts: Write C# code using Aspose.Cells to apply a custom three‑icon set to a specified cell range and save the workbook. | Explain how to modify icon order or replace the built‑in TrafficLights31 set with custom images in Aspose.Cells conditional formatting. | Show how to add multiple icon‑set rules to different ranges within the same worksheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsIconSetExample
{
    // Creates a workbook, fills cells A1:A10 with incremental sales values, defines a conditional‑formatting rule, adds the built‑in TrafficLights31 icon set, optionally displays cell values, keeps the default green‑yellow‑red order, customizes each icon's type and index, and saves the result as SalesIconSet.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales figures in column A (rows 1‑10)
            for (int i = 0; i < 10; i++)
            {
                // Example values ranging from low to high
                sheet.Cells[i, 0].PutValue((i + 1) * 10);
            }

            // Get the conditional formatting collection of the worksheet
            ConditionalFormattingCollection cfCollection = sheet.ConditionalFormattings;

            // Add a new conditional formatting rule container
            int cfIndex = cfCollection.Add();
            FormatConditionCollection fcCollection = cfCollection[cfIndex];

            // Define the cell area (A1:A10) to which the icon set will be applied
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcCollection.AddArea(area);

            // Add an IconSet condition
            int conditionIndex = fcCollection.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcCollection[conditionIndex];

            // Use the built‑in TrafficLights31 set (green, yellow, red)
            condition.IconSet.Type = IconSetType.TrafficLights31;
            condition.IconSet.ShowValue = true;   // optional: display cell values alongside icons
            condition.IconSet.Reverse = false;    // keep default order (green → yellow → red)

            // Customize individual icons if needed (here we explicitly set each icon's type and index)
            ConditionalFormattingIcon icon0 = condition.IconSet.CfIcons[0];
            icon0.Type = IconSetType.TrafficLights31; // green icon
            icon0.Index = 0;

            ConditionalFormattingIcon icon1 = condition.IconSet.CfIcons[1];
            icon1.Type = IconSetType.TrafficLights31; // yellow icon
            icon1.Index = 1;

            ConditionalFormattingIcon icon2 = condition.IconSet.CfIcons[2];
            icon2.Type = IconSetType.TrafficLights31; // red icon
            icon2.Index = 2;

            // Save the workbook with the applied conditional icon set
            workbook.Save("SalesIconSet.xlsx");
        }
    }
}
