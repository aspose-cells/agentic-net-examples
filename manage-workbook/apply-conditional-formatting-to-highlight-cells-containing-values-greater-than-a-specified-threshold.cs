// Title: C# Aspose.Cells: Highlight Cells Greater Than a Threshold with Conditional Formatting
// Description: Creates a new workbook, fills A1:A10 with values 0‑90, sets a threshold of 50, and adds a conditional formatting rule that colors cells orange with black font when their value exceeds the threshold. The workbook is saved as ConditionalFormattingGreaterThan.xlsx.
// Keywords: Aspose.Cells conditional formatting C# | highlight cells greater than threshold | Aspose.Cells .NET example | conditional formatting orange background | cell value greater than operator | C# Excel automation Aspose
// Common Searches: Aspose.Cells how to apply conditional formatting for values above a limit | C# conditional formatting greater than threshold Aspose.Cells | set cell background color based on value Aspose.Cells .NET | example of conditional formatting range A1:A10 Aspose
// Developer Intent: Generate a .NET workbook and apply a conditional formatting rule that highlights cells whose numeric value is greater than a specified threshold.
// Use Cases: Flag sales figures that exceed a target amount in a financial dashboard. | Mark test scores above the passing grade for quick visual review. | Identify inventory levels that surpass a reorder limit.
// AI Prompts: Write C# code with Aspose.Cells that colors cells red when values are below a minimum threshold. | Show how to add multiple conditional formatting rules to different ranges in the same worksheet using Aspose.Cells. | Demonstrate changing the font style of cells that meet a conditional formatting condition in Aspose.Cells for .NET.

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, fills A1:A10 with values 0‑90, sets a threshold of 50, and adds a conditional formatting rule that colors cells orange with black font when their value exceeds the threshold. The workbook is saved as ConditionalFormattingGreaterThan.xlsx.
class ConditionalFormattingExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
        }

        // Define the threshold value
        double threshold = 50;

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Define the range to which the conditional formatting will be applied (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add a condition: highlight cells where the value is greater than the threshold
        int conditionIndex = fcc.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            threshold.ToString(),
            null);
        FormatCondition fc = fcc[conditionIndex];

        // Set the formatting style for cells that meet the condition
        fc.Style.BackgroundColor = Color.Orange;
        fc.Style.Font.Color = Color.Black;

        // Save the workbook
        workbook.Save("ConditionalFormattingGreaterThan.xlsx");
    }
}
