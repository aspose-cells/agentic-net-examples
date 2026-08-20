// Title: Aspose.Cells C# – Apply Conditional Formatting Using a Dynamic Threshold Formula
// Description: Creates a workbook, fills A1:A10 with numbers, sets B1 to AVERAGE(A1:A10)+5, adds a conditional formatting rule that colors cells in A1:A10 yellow when their value exceeds the dynamic threshold in B1, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | dynamic threshold | formula reference | cell value condition | FormatCondition | OperatorType.GreaterThan | Excel automation | highlight cells | average plus offset
// Common Searches: Aspose.Cells conditional formatting with formula reference | C# apply conditional formatting based on another cell | dynamic threshold conditional formatting .NET | highlight cells greater than average using Aspose.Cells | how to use cell B1 as threshold in Aspose.Cells
// Developer Intent: Create a conditional formatting rule that highlights cells when their value is greater than a threshold calculated by a formula in another cell.
// Use Cases: Mark outliers in a data column that exceed the column average plus a safety margin. | Flag expense entries higher than the average expense plus a configurable buffer. | Visualize sensor readings that surpass a dynamically computed safety limit.
// AI Prompts: Generate C# code with Aspose.Cells to conditionally format column D based on a formula in cell E1. | Show how to set a dynamic threshold using AVERAGE and an offset for conditional formatting across multiple ranges in Aspose.Cells. | Provide an example of referencing $B$1 in a conditional formatting rule to compare cell values in Aspose.Cells.

using System;
using Aspose.Cells;
using System.Drawing;

namespace ConditionalFormattingDemo
{
    // Creates a workbook, fills A1:A10 with numbers, sets B1 to AVERAGE(A1:A10)+5, adds a conditional formatting rule that colors cells in A1:A10 yellow when their value exceeds the dynamic threshold in B1, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data in column A (rows 1-10)
            for (int i = 0; i < 10; i++)
            {
                worksheet.Cells[i, 0].PutValue(i * 10 + 5); // Example values: 5,15,25,...
            }

            // Define a dynamic threshold in cell B1 using a formula (e.g., average + 5)
            worksheet.Cells["B1"].Formula = "AVERAGE(A1:A10)+5";

            // Add a conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

            // Set the range to which the conditional formatting will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add a CellValue condition: highlight cells where value > threshold (B1)
            // Use the AddCondition method with parameters (type, operator, formula1, formula2)
            int conditionIndex = fcs.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "=B$1",   // Formula1 refers to the dynamic threshold cell
                null);    // No second formula needed for GreaterThan

            // Retrieve the created condition and set its formatting style
            FormatCondition fc = fcs[conditionIndex];
            fc.Style.BackgroundColor = Color.Yellow;

            // Save the workbook
            workbook.Save("ConditionalFormattingDynamicThreshold.xlsx");
        }
    }
}
