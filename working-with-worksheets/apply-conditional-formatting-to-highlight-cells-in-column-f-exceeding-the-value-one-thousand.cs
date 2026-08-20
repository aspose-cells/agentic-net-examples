// Title: Aspose.Cells for .NET C# – Highlight Column F Cells Over 1000 with Conditional Formatting
// Description: Creates a workbook, defines a CellArea for column F, adds a CellValue rule (value > 1000), sets a yellow background style, and saves the file as ColumnF_ConditionalFormatting.xlsx.
// Keywords: Aspose.Cells | C# | conditional formatting | column F | value greater than 1000 | highlight cells | background color | Excel automation | FormatCondition | GitHub sample
// Common Searches: Aspose.Cells conditional formatting column F | C# highlight cells greater than 1000 | set background color based on value Aspose.Cells | add CellValue rule Aspose.Cells .NET | Excel conditional formatting example C#
// Developer Intent: Add a conditional formatting rule that colors cells in column F yellow when their numeric value exceeds 1000.
// Use Cases: Financial statements where expenses over 1,000 in column F are automatically flagged. | Data‑analysis dashboards that draw attention to outlier values in a specific column. | Reusable spreadsheet templates that visually emphasize high‑value entries without manual formatting.
// AI Prompts: Generate C# code using Aspose.Cells to apply a yellow background to column F cells whose value is greater than 1000. | Show how to extend the snippet to apply the same conditional formatting to multiple columns or a dynamic row range. | Explain how to replace the CellValue condition with a formula‑based rule to highlight column F values exceeding 1000.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, defines a CellArea for column F, adds a CellValue rule (value > 1000), sets a yellow background style, and saves the file as ColumnF_ConditionalFormatting.xlsx.
class ConditionalFormattingExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range for column F (zero‑based column index 5)
        CellArea range = new CellArea
        {
            StartRow = 0,      // first row (A1)
            EndRow = 100,      // adjust as needed
            StartColumn = 5,   // column F
            EndColumn = 5
        };

        // Add a CellValue condition: value > 1000
        // The Add method returns an array where the first element is the condition index
        int[] addResult = fcc.Add(
            range,
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            "1000",   // Formula1 – the threshold value
            null);    // Formula2 – not required for GreaterThan

        // Retrieve the created condition and set its style (e.g., yellow background)
        FormatCondition condition = fcc[addResult[0]];
        condition.Style.BackgroundColor = Color.Yellow;

        // Save the workbook
        workbook.Save("ColumnF_ConditionalFormatting.xlsx");
    }
}
