// Title: Aspose.Cells .NET: Apply Theme Accent2 Fill via Conditional Formatting
// Description: Creates a workbook, populates A1:A10 with values 0‑90, adds a conditional format that highlights cells > 50 using the workbook’s Accent2 theme color as a solid fill, and saves the file as ConditionalFormattingAccent2.xlsx.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | theme color | Accent2 | solid fill | Style.ForegroundColor | ThemeColorType | Excel automation
// Common Searches: Aspose.Cells set conditional format theme color | C# apply Accent2 fill in Excel using Aspose | how to use workbook theme colors in conditional formatting .NET | conditional formatting solid fill based on value Aspose.Cells
// Developer Intent: Add a conditional formatting rule that fills cells with the workbook’s Accent2 theme color when their numeric value exceeds a threshold.
// Use Cases: Brand‑consistent highlighting of KPI values that surpass a target. | Automatic coloring of financial figures above a limit using the document’s theme. | Applying uniform theme‑based conditional styles across multiple sheets in a reporting workbook.
// AI Prompts: Generate code to use the Accent3 theme color instead of Accent2 for the same condition. | Show how to apply the Accent2 fill rule to range B2:B20 with a "less than 30" condition. | Explain how to retrieve the RGB value of the Accent2 theme color after the workbook is saved.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, populates A1:A10 with values 0‑90, adds a conditional format that highlights cells > 50 using the workbook’s Accent2 theme color as a solid fill, and saves the file as ConditionalFormattingAccent2.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in column A (rows 1-10)
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
        }

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range to which the conditional formatting will be applied (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add a condition: cells with value greater than 50
        int conditionIndex = fcc.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            "50",
            null);
        FormatCondition condition = fcc[conditionIndex];

        // Create a style that uses the workbook's theme Accent2 color for the fill
        Style style = workbook.CreateStyle();

        // Create a CellsColor instance and set its ThemeColor to Accent2 (no tint)
        CellsColor themeColor = workbook.CreateCellsColor();
        themeColor.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);

        // Apply the theme color to the style's foreground (fill) and set solid pattern
        style.ForegroundColor = themeColor.Color;
        style.Pattern = BackgroundType.Solid;

        // Assign the style to the conditional format
        condition.Style = style;

        // Save the workbook
        workbook.Save("ConditionalFormattingAccent2.xlsx");
    }
}
