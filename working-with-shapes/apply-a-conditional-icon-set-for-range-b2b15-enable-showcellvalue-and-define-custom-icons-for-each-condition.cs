// Title: C# – Apply Custom Icon Set Conditional Formatting with ShowValue to B2:B15 using Aspose.Cells
// Description: Creates a workbook, populates cells B2‑B15 with incremental numbers, adds an IconSet conditional format (Arrows3) to the range, enables ShowValue so the numeric values appear beside the icons, customizes each of the three icons with different built‑in types, and saves the file.
// Keywords: Aspose.Cells | C# | IconSet conditional formatting | ShowValue | custom icons | ConditionalFormattingIcon | Excel icon set | .NET | cell range B2:B15 | FormatConditionType.IconSet
// Common Searches: Aspose.Cells add icon set to range | Enable ShowValue for icon set in Aspose.Cells .NET | Customize icons in Aspose.Cells conditional formatting | C# code for IconSet conditional format | How to change icon types in Aspose.Cells
// Developer Intent: Add an IconSet conditional format to B2:B15, display cell values, and assign distinct built‑in icons to each condition.
// Use Cases: Show sales trends with green, yellow, and red arrows while keeping the actual figures visible. | Mark project milestones using different box icons for low, medium, and high completion percentages. | Highlight financial risk levels by applying custom arrow and box icons to a column of values.
// AI Prompts: Generate C# code with Aspose.Cells to apply a 4‑icon set to range C3:C20, hide the cell values, and assign specific icons for each threshold. | Provide an example of using Aspose.Cells to create an icon set where each icon is a custom image file instead of a built‑in icon. | Show how to switch the IconSet type at runtime based on user input while preserving the ShowValue setting in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, populates cells B2‑B15 with incremental numbers, adds an IconSet conditional format (Arrows3) to the range, enables ShowValue so the numeric values appear beside the icons, customizes each of the three icons with different built‑in types, and saves the file.
class IconSetConditionalFormattingExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the range B2:B15 (rows 1-14, column 1)
        for (int row = 1; row <= 14; row++)
        {
            sheet.Cells[row, 1].PutValue(row * 10); // Example values: 10, 20, ..., 140
        }

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the cell area B2:B15
        CellArea area = new CellArea
        {
            StartRow = 1,   // B2 -> row index 1
            EndRow = 14,    // B15 -> row index 14
            StartColumn = 1, // Column B -> index 1
            EndColumn = 1
        };
        fcs.AddArea(area);

        // Add an IconSet condition
        int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcs[conditionIdx];

        // Set the icon set type (using a 3‑icon set as an example)
        condition.IconSet.Type = IconSetType.Arrows3;

        // Ensure the cell values are displayed alongside the icons
        condition.IconSet.ShowValue = true;

        // Customize each icon in the set
        // Icon 0
        ConditionalFormattingIcon icon0 = condition.IconSet.CfIcons[0];
        icon0.Type = IconSetType.Arrows3; // Use the standard arrows icon
        icon0.Index = 0;                  // Position in the set

        // Icon 1
        ConditionalFormattingIcon icon1 = condition.IconSet.CfIcons[1];
        icon1.Type = IconSetType.ArrowsGray3; // Gray arrows for the middle condition
        icon1.Index = 1;

        // Icon 2
        ConditionalFormattingIcon icon2 = condition.IconSet.CfIcons[2];
        icon2.Type = IconSetType.Boxes5; // Boxes icon for the highest condition
        icon2.Index = 2;

        // Save the workbook
        workbook.Save("IconSetConditionalFormatting_B2_B15.xlsx");
    }
}
