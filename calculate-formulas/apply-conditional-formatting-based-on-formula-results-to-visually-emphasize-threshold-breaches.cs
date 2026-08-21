// Title: C# – Apply Formula‑Based Conditional Formatting in Aspose.Cells to Highlight Values Over 100
// Description: Creates a new workbook, fills A1:A10 with incremental numbers, defines the range, adds an expression‑based conditional formatting rule "=A1>100", and formats matching cells with a yellow background, red bold font, then saves the file.
// Keywords: Aspose.Cells | conditional formatting | formula based formatting | expression condition | C# | .NET Excel automation | highlight cells over threshold | cell background color | bold red font | Excel workbook generation | range formatting Aspose
// Common Searches: Aspose.Cells conditional formatting example C# | how to highlight cells greater than 100 using Aspose.Cells | apply expression based conditional formatting Aspose.Cells .NET | set cell background color based on value Aspose.Cells | conditional formatting range C# Aspose.Cells
// Developer Intent: Add a conditional formatting rule that colors cells yellow and makes the font red‑bold when the cell value exceeds 100.
// Use Cases: Flag sales figures that surpass a target in automated reports. | Mark temperature readings above safety limits in engineering worksheets. | Highlight overdue days in project timelines when they exceed a deadline.
// AI Prompts: Generate C# code with Aspose.Cells that applies green background for values < 50 and red background for values > 150. | Show how to modify the example to use a dynamic range and reference the active cell in the expression. | Explain steps to export the workbook containing conditional formatting to PDF while preserving styles.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingDemo
{
    // Creates a new workbook, fills A1:A10 with incremental numbers, defines the range, adds an expression‑based conditional formatting rule "=A1>100", and formats matching cells with a yellow background, red bold font, then saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (rows 0-9)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i * 20); // Values: 0,20,40,...,180
            }

            // Define the range to which the conditional formatting will be applied (A1:A10)
            CellArea range = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Add the target range to the collection
            fcc.AddArea(range);

            // Add a condition that highlights cells where the value exceeds 100
            // Using Expression type with a formula that evaluates to TRUE when the condition is met
            int[] result = fcc.Add(
                range,
                FormatConditionType.Expression,
                OperatorType.None,
                "=A1>100",
                null);

            // Retrieve the created FormatCondition (index is result[0])
            FormatCondition condition = fcc[result[0]];

            // Set the visual style for cells that meet the condition
            condition.Style.BackgroundColor = Color.Yellow;
            condition.Style.Font.Color = Color.Red;
            condition.Style.Font.IsBold = true;

            // Save the workbook
            workbook.Save("ThresholdConditionalFormatting.xlsx");
        }
    }
}
