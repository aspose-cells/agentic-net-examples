// Title: Apply Green Background Conditional Formatting for Values >1000 with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, defines the range A1:A20, adds a conditional formatting rule, and sets a green fill for any cell whose numeric value is greater than 1000. The workbook is saved as ConditionalFormatting_GreaterThan1000.xlsx.
// Keywords: Aspose.Cells | C# conditional formatting | green background | value greater than 1000 | .NET spreadsheet | FormatCondition | CellArea | Excel conditional formatting example
// Common Searches: Aspose.Cells conditional formatting example C# | how to set cell background color based on value Aspose.Cells | C# highlight cells greater than 1000 in Excel | apply green fill when value exceeds threshold Aspose.Cells | conditional formatting range A1:A20 Aspose.Cells
// Developer Intent: Create a conditional formatting rule that colors cells green when their numeric value exceeds 1000.
// Use Cases: Flag high‑value sales figures in financial dashboards. | Mark budget overruns in expense tracking sheets. | Highlight safety‑critical measurements that surpass a defined limit.
// AI Prompts: Generate C# code using Aspose.Cells to apply a red background to cells in column B when the value is below 200. | Show how to add multiple conditional formatting rules with different colors for low, medium, and high value ranges in a worksheet. | Explain how to replace the constant "1000" with a cell‑referenced formula in an Aspose.Cells conditional formatting rule.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingExample
{
    // C# example that creates a workbook, defines the range A1:A20, adds a conditional formatting rule, and sets a green fill for any cell whose numeric value is greater than 1000. The workbook is saved as ConditionalFormatting_GreaterThan1000.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to which the conditional formatting will be applied (e.g., A1:A20)
            CellArea range = new CellArea
            {
                StartRow = 0,    // Row 1 (zero‑based index)
                EndRow = 19,     // Row 20
                StartColumn = 0, // Column A
                EndColumn = 0    // Column A
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

            // Associate the defined range with the conditional formatting collection
            conditions.AddArea(range);

            // Add a condition: cell value greater than 1000
            int conditionIndex = conditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "1000",   // Formula1 – the threshold value
                null);    // Formula2 – not needed for GreaterThan

            // Retrieve the created condition and set its style (green background)
            FormatCondition condition = conditions[conditionIndex];
            condition.Style.BackgroundColor = Color.Green;

            // (Optional) Populate some sample data to demonstrate the rule
            for (int i = 0; i < 20; i++)
            {
                worksheet.Cells[i, 0].PutValue(i * 150); // Values: 0,150,300,...,2850
            }

            // Save the workbook
            workbook.Save("ConditionalFormatting_GreaterThan1000.xlsx");
        }
    }
}
