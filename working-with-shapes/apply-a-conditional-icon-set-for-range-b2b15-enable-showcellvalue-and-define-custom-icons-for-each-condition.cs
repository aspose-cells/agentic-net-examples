// Title: Aspose.Cells .NET – Apply a Custom Icon Set with ShowValue to Range B2:B15
// Description: Creates a workbook, defines the B2:B15 range, adds an IconSet conditional format, enables ShowValue so cell values appear beside icons, and replaces the three default icons with custom types before saving the file.
// Keywords: Aspose.Cells conditional formatting | icon set .NET | ShowValue property | custom icons Aspose.Cells | C# Excel icon set example | modify IconSet icons | conditional formatting icons
// Common Searches: Aspose.Cells add icon set to range | Enable ShowValue for icon set conditional formatting | Change individual icons in Aspose.Cells IconSet | C# example conditional formatting with icons | Replace default icons in Aspose.Cells
// Developer Intent: Add an IconSet conditional format to cells B2:B15, display the numeric values alongside the icons, and assign a specific icon type to each position in the set.
// Use Cases: Display traffic‑light indicators for KPI scores while keeping the numeric value visible. | Show upward/downward trend arrows with custom colors in a financial dashboard. | Mark project phases with distinct box icons next to percentage completions.
// AI Prompts: Generate C# code using Aspose.Cells to apply a three‑icon set to B2:B15, turn on ShowValue, and set custom IconSetType values for each icon. | Explain how to modify the ShowValue flag of an IconSet conditional format in Aspose.Cells for .NET. | Provide step‑by‑step instructions to replace default icons in an existing Aspose.Cells IconSet rule with specific icon types.

using System;
using Aspose.Cells;

// Creates a workbook, defines the B2:B15 range, adds an IconSet conditional format, enables ShowValue so cell values appear beside icons, and replaces the three default icons with custom types before saving the file.
class IconSetConditionalFormattingExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range B2:B15 (rows 1‑14, column 1)
        CellArea area = new CellArea
        {
            StartRow = 1,   // B2
            EndRow = 14,    // B15
            StartColumn = 1,
            EndColumn = 1
        };
        fcs.AddArea(area);

        // Add an IconSet condition
        int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcs[conditionIdx];

        // Set the icon set type (using a 3‑icon set as a base)
        condition.IconSet.Type = IconSetType.Arrows3;

        // Ensure cell values are shown alongside icons
        condition.IconSet.ShowValue = true;

        // Customize each icon in the set
        // First icon
        ConditionalFormattingIcon icon0 = condition.IconSet.CfIcons[0];
        icon0.Type = IconSetType.Arrows3;   // Arrow icon
        icon0.Index = 0;                    // Position in the set

        // Second icon
        ConditionalFormattingIcon icon1 = condition.IconSet.CfIcons[1];
        icon1.Type = IconSetType.ArrowsGray3; // Gray arrow icon
        icon1.Index = 1;

        // Third icon
        ConditionalFormattingIcon icon2 = condition.IconSet.CfIcons[2];
        icon2.Type = IconSetType.Boxes5;   // Box icon
        icon2.Index = 2;

        // (Optional) Populate some sample data in the range
        for (int row = 1; row <= 14; row++)
        {
            sheet.Cells[row, 1].PutValue(row * 10); // B2 = 20, B3 = 30, ...
        }

        // Save the workbook
        workbook.Save("IconSetConditionalFormatting.xlsx");
    }
}
