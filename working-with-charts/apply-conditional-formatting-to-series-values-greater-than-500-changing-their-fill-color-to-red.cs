// Title: Aspose.Cells for .NET – Conditional Formatting to Highlight Cells > 500 with Red Fill
// Description: Creates a workbook, optionally fills column A with sample numbers, defines a conditional‑formatting range (A1:A10), adds a rule that selects cells whose value exceeds 500, and applies a red background style before saving the file.
// Keywords: Aspose.Cells | .NET | C# conditional formatting | highlight cells greater than 500 | red fill color | Excel automation | FormatCondition | CellArea range | background color rule
// Common Searches: Aspose.Cells highlight cells >500 red | C# conditional formatting Excel Aspose | set cell background based on value Aspose.Cells | apply red fill to values over 500 using .NET | how to add conditional formatting with Aspose.Cells
// Developer Intent: Add a conditional‑formatting rule that colors any cell with a numeric value above 500 red.
// Use Cases: Flag sales numbers that exceed a target in a financial dashboard. | Mark temperature readings above safety limits in an engineering log. | Identify budget items that surpass allocated amounts in a planning sheet.
// AI Prompts: Write C# code using Aspose.Cells to apply a red fill to cells greater than 500 in a specified range. | Show how to combine multiple conditional‑formatting rules (e.g., >500 red, <200 green) on the same worksheet with Aspose.Cells. | Explain how to change the fill color of an existing conditional‑formatting rule programmatically in Aspose.Cells for .NET.

using Aspose.Cells;
using System.Drawing;

// Creates a workbook, optionally fills column A with sample numbers, defines a conditional‑formatting range (A1:A10), adds a rule that selects cells whose value exceeds 500, and applies a red background style before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample series data in column A (optional)
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i * 150); // 0, 150, 300, ... 1350
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

        // Add a condition: cells with values greater than 500
        int conditionIndex = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "500", null);
        FormatCondition condition = fcc[conditionIndex];

        // Set the fill color to red for cells that meet the condition
        condition.Style.BackgroundColor = Color.Red;

        // Save the workbook
        workbook.Save("ConditionalFormattingGreaterThan500.xlsx");
    }
}
