// Title: C# – Apply Theme Accent2 Fill via Conditional Formatting in Aspose.Cells
// Description: Creates a workbook, fills A1:A10 with numeric values, adds a Between (30‑70) conditional formatting rule, and sets the cell background to the workbook’s Accent2 theme color using a solid fill before saving the file.
// Keywords: Aspose.Cells | C# | conditional formatting | theme color | Accent2 | solid fill | cell style | Excel automation | highlight values | between condition
// Common Searches: Aspose.Cells set conditional formatting theme color | C# conditional formatting Accent2 fill | how to use theme colors in Aspose.Cells | apply solid background based on value range Aspose.Cells | conditional formatting between values C#
// Developer Intent: Create a conditional formatting rule that fills cells with the workbook’s Accent2 theme color when their values are between 30 and 70.
// Use Cases: Highlight sales figures that fall within a target range using the corporate Accent2 color for brand consistency. | Mark expense entries between two thresholds in a financial report with a theme‑based fill to draw attention. | Apply a uniform Accent2 background to multiple worksheets that share the same value‑range criteria, ensuring visual consistency across the workbook.
// AI Prompts: Show how to change the conditional formatting to use the Accent3 theme color instead of Accent2. | Provide code to add a second rule that colors cells outside the 30‑70 range with a different theme color. | Explain how to apply the same Accent2 conditional formatting to non‑contiguous ranges such as A1:A10 and C1:C10.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, fills A1:A10 with numeric values, adds a Between (30‑70) conditional formatting rule, and sets the cell background to the workbook’s Accent2 theme color using a solid fill before saving the file.
class ApplyAccent2ConditionalFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
        }

        // Define the range to which the conditional formatting will be applied (A1:A10)
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };

        // Add a new conditional formatting collection
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];
        fcc.AddArea(range);

        // Add a condition: highlight cells with values between 30 and 70 (inclusive)
        int conditionIdx = fcc.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.Between,
            "30",
            "70");

        FormatCondition condition = fcc[conditionIdx];

        // Create a CellsColor that uses the theme's Accent2 color (no tint)
        CellsColor accent2Color = workbook.CreateCellsColor();
        accent2Color.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);

        // Create a style that uses the Accent2 color as the fill background
        Style style = workbook.CreateStyle();
        style.ForegroundColor = accent2Color.Color; // Apply the theme color
        style.Pattern = BackgroundType.Solid;       // Solid fill

        // Assign the style to the conditional format
        condition.Style = style;

        // Save the workbook
        workbook.Save("Accent2ConditionalFormatting.xlsx");
    }
}
