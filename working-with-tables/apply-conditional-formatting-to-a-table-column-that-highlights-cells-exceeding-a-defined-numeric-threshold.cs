// Title: Aspose.Cells .NET – Conditional Formatting to Highlight Column Cells Over a Threshold
// Description: Creates a workbook, fills column B with incremental numbers, adds a conditional‑formatting rule for rows 0‑19 that colors cells yellow when the value exceeds 50, and saves the file as ColumnConditionalFormatting.xlsx.
// Keywords: Aspose.Cells conditional formatting C# | highlight cells greater than value .NET | column threshold formatting Aspose | yellow background conditional rule | C# Excel cell value condition
// Common Searches: Aspose.Cells highlight column values above a limit | C# conditional formatting for a specific column in Excel | apply yellow background when cell > 50 using Aspose.Cells | add cell‑value rule to column B in Aspose.Cells workbook | how to set conditional format range in Aspose.Cells .NET
// Developer Intent: Add a conditional‑formatting rule that colors cells in a column yellow when their numeric value is greater than a defined threshold.
// Use Cases: Mark sales figures above target in a financial report. | Flag temperature readings that exceed safety limits. | Identify project tasks with overdue days greater than a set number.
// AI Prompts: Generate C# Aspose.Cells code that applies a red background to cells in column C when values are less than 20. | Show how to create multiple conditional‑formatting rules for different columns in the same worksheet using Aspose.Cells .NET. | Provide an example of a formula‑based conditional format that highlights rows where the sum of two columns exceeds a specified threshold.

using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingExample
{
    // Creates a workbook, fills column B with incremental numbers, adds a conditional‑formatting rule for rows 0‑19 that colors cells yellow when the value exceeds 50, and saves the file as ColumnConditionalFormatting.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample numeric data in column B (index 1).
            for (int row = 0; row < 20; row++)
            {
                worksheet.Cells[row, 1].PutValue(row * 10); // Values: 0,10,20,...
            }

            // Add a new conditional formatting collection to the worksheet.
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

            // Define the range for the entire column B (rows 0‑19).
            CellArea columnRange = new CellArea
            {
                StartRow = 0,
                EndRow = 19,
                StartColumn = 1,
                EndColumn = 1
            };
            fcc.AddArea(columnRange);

            // Add a condition that highlights cells with values greater than the threshold (e.g., 50).
            int conditionIndex = fcc.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",          // Threshold value as a string.
                null);         // No second formula needed for GreaterThan.

            // Retrieve the created condition and set its formatting style.
            FormatCondition condition = fcc[conditionIndex];
            condition.Style.BackgroundColor = Color.Yellow; // Highlight color.

            // Save the workbook to a file.
            workbook.Save("ColumnConditionalFormatting.xlsx");
        }
    }
}
