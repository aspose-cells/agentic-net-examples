// Title: C# Aspose.Cells – Conditional Formatting to Color Cells Red When Value > 500
// Description: Demonstrates how to create a workbook, populate column A with numeric data, define a conditional‑formatting rule for the range A1:A10, and set a red background for cells whose value exceeds 500 using Aspose.Cells for .NET.
// Keywords: Aspose.Cells conditional formatting C# | highlight cells >500 red | set cell background color Aspose.Cells | C# workbook conditional rule | format cells based on value Aspose | SeriesConditionalFormatting.xlsx example | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells apply conditional formatting C# | color cells red when value greater than 500 | conditional formatting range A1:A10 Aspose | how to set background color based on cell value Aspose.Cells | C# example conditional formatting workbook
// Developer Intent: Add a rule that turns cells red when their numeric value is higher than 500.
// Use Cases: Highlight sales numbers that exceed a target threshold. | Flag sensor readings above safe limits for quick review. | Mark budget items that surpass allocated amounts.
// AI Prompts: Generate C# Aspose.Cells code that colors cells green when the value is below 200. | Show how to add multiple conditional formatting rules (red >500, yellow 300‑500) to the same range. | Provide an example of applying conditional formatting to a chart data series instead of worksheet cells in Aspose.Cells for .NET.

using Aspose.Cells;
using System.Drawing;

// Demonstrates how to create a workbook, populate column A with numeric data, define a conditional‑formatting rule for the range A1:A10, and set a red background for cells whose value exceeds 500 using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Populate sample series values in column A
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
        workbook.Save("SeriesConditionalFormatting.xlsx");
    }
}
