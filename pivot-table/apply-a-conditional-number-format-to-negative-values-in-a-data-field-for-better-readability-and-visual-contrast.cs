// Title: C# Aspose.Cells – Apply Conditional Number Format to Negative Values in a Column
// Description: Demonstrates how to create a workbook, fill cells A1:A5 with mixed numbers, add a conditional formatting rule for values less than zero, define a custom style "[Red]-0;[Red]-0;0" to show negatives in red with a minus sign, and save the file as an XLSX document.
// Keywords: Aspose.Cells conditional formatting C# | custom number format negative values | red negative numbers Excel | conditional style negative cells .NET | apply number format Aspose.Cells
// Common Searches: Aspose.Cells set custom format for negative numbers | C# conditional formatting red negative values Excel | how to highlight negative cells with Aspose.Cells | apply number format to a range using Aspose.Cells
// Developer Intent: Generate an Excel workbook where any cell with a value below zero is displayed in red using a custom number format.
// Use Cases: Mark financial losses in generated reports with red formatting. | Flag inventory shortages for quick visual identification. | Display negative account balances distinctly in accounting sheets.
// AI Prompts: Write C# code with Aspose.Cells that formats numbers less than zero in red for the range B2:B15 and saves the workbook. | Create a reusable method that adds a conditional formatting rule to any worksheet to apply a custom style to negative values. | Explain how to modify the custom number format string to show negative numbers in parentheses instead of a minus sign.

using System;
using Aspose.Cells;

namespace AsposeCellsConditionalNumberFormat
{
    // Demonstrates how to create a workbook, fill cells A1:A5 with mixed numbers, add a conditional formatting rule for values less than zero, define a custom style "[Red]-0;[Red]-0;0" to show negatives in red with a minus sign, and save the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with both positive and negative values
                sheet.Cells["A1"].PutValue(150);
                sheet.Cells["A2"].PutValue(-75);
                sheet.Cells["A3"].PutValue(200);
                sheet.Cells["A4"].PutValue(-30);
                sheet.Cells["A5"].PutValue(0);

                // Add a conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

                // Define the range to which the conditional formatting will be applied (A1:A5)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 4,
                    StartColumn = 0,
                    EndColumn = 0
                };
                cfCollection.AddArea(area);

                // Add a condition for cells with values less than 0 (negative numbers)
                int conditionIndex = cfCollection.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.LessThan,
                    "0",
                    null);

                // Retrieve the created condition
                FormatCondition condition = cfCollection[conditionIndex];

                // Create a style that defines a custom number format for negative values
                // The format displays negative numbers in red with a minus sign
                Style negativeStyle = workbook.CreateStyle();
                negativeStyle.Custom = "[Red]-0;[Red]-0;0";

                // Apply the style to the condition
                condition.Style = negativeStyle;

                // Save the workbook
                string outputPath = "ConditionalNegativeNumberFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
