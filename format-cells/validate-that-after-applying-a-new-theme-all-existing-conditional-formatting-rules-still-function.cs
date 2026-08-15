// Title: C# Example: Verify Conditional Formatting Persists After Applying a Custom Theme with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a conditional formatting rule that colors cells red when the value exceeds 15, apply a custom theme via Workbook.CustomTheme, and programmatically confirm that the rule still works using GetConditionalFormattingResult before saving the file.
// Keywords: Aspose.Cells | .NET | C# | custom theme | Workbook.CustomTheme | conditional formatting | validation | GetConditionalFormattingResult | preserve formatting | sample code
// Common Searches: Aspose.Cells verify conditional formatting after theme change | C# check conditional formatting persistence with custom theme | GetConditionalFormattingResult after Workbook.CustomTheme | how to test conditional formatting after applying a theme in Aspose.Cells | Aspose.Cells .NET theme validation example
// Developer Intent: Confirm that existing conditional formatting rules remain effective after a custom theme is applied to a workbook.
// Use Cases: Generate a styled report where theme colors are customized but conditional formatting must stay unchanged. | Automate a unit test that applies Workbook.CustomTheme to a workbook with predefined conditional formatting and asserts the expected font color. | Create a reusable component that validates conditional formatting after any theme modification in an Aspose.Cells .NET project.
// AI Prompts: Write C# code using Aspose.Cells to apply a custom theme and then verify that a conditional formatting rule still changes the font color to red for values greater than 15. | Explain the role of GetConditionalFormattingResult after a theme change and how to interpret its ConditionalStyle properties. | Provide a unit‑test snippet that checks conditional formatting persistence after calling Workbook.CustomTheme in an Aspose.Cells .NET application.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeValidation
{
    // Demonstrates how to create a workbook, add a conditional formatting rule that colors cells red when the value exceeds 15, apply a custom theme via Workbook.CustomTheme, and programmatically confirm that the rule still works using GetConditionalFormattingResult before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);

            // Add a conditional formatting that highlights values > 15 with red font
            int cfIndex = sheet.ConditionalFormattings.Add();                         // Add empty conditional formatting
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];   // Retrieve the collection

            // Define the range A1:A4
            CellArea area = new CellArea { StartRow = 0, EndRow = 3, StartColumn = 0, EndColumn = 0 };
            fcc.AddArea(area);

            // Add the condition (rule) – using AddCondition rule
            int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "15", null);
            FormatCondition fc = fcc[conditionIdx];
            // Set the style for the condition
            Style redStyle = workbook.CreateStyle();
            redStyle.Font.Color = Color.Red;
            fc.Style = redStyle;

            // Apply a custom theme (custom theme rule)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(50, 50, 50),    // Text2
                Color.FromArgb(255, 0, 0),     // Accent1
                Color.FromArgb(0, 255, 0),     // Accent2
                Color.FromArgb(0, 0, 255),     // Accent3
                Color.FromArgb(255, 255, 0),   // Accent4
                Color.FromArgb(255, 0, 255),   // Accent5
                Color.FromArgb(0, 255, 255),   // Accent6
                Color.FromArgb(0, 0, 255),     // Hyperlink
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink
            };
            workbook.CustomTheme("ValidationTheme", customColors); // Apply the theme

            // Validate that conditional formatting still works after theme change
            Console.WriteLine("Conditional Formatting Validation After Theme Application:");
            for (int row = 0; row <= 3; row++)
            {
                Cell cell = sheet.Cells[row, 0]; // Column A
                ConditionalFormattingResult result = cell.GetConditionalFormattingResult(); // Get result rule
                bool isFormatted = result?.ConditionalStyle != null;
                string formattedInfo = isFormatted ? $"Font Color = {result.ConditionalStyle.Font.Color}" : "No formatting applied";
                Console.WriteLine($"Cell {cell.Name}: Value = {cell.Value}, {formattedInfo}");
            }

            // Save the workbook (save rule)
            workbook.Save("ThemeValidationResult.xlsx");
        }
    }
}
