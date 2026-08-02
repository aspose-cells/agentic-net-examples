// Title: C# – Ensure Conditional Formatting Survives Theme Changes in Aspose.Cells
// Description: The sample creates a workbook, adds a between‑value conditional format to cells A1:A5, applies a 12‑color custom theme with Workbook.CustomTheme, then iterates through the range using GetConditionalFormattingResult to confirm the rule still triggers and the red background is retained before saving the file.
// Keywords: Aspose.Cells | C# | conditional formatting | custom theme | Workbook.CustomTheme | GetConditionalFormattingResult | validation | between values | Excel styling programmatically | unit test
// Common Searches: Aspose.Cells verify conditional formatting after theme change | C# check conditional format persistence with custom theme | GetConditionalFormattingResult example Aspose.Cells | how to test conditional formatting after applying Workbook.CustomTheme | conditional formatting unit test Aspose.Cells .NET
// Developer Intent: Confirm that conditional formatting rules continue to evaluate correctly and display the defined style after a workbook's theme is modified.
// Use Cases: Apply a corporate color scheme to a generated report while preserving threshold‑based highlights. | Automated regression test that changes the workbook theme and asserts conditional styles remain unchanged. | Dynamic theming of dashboards where existing conditional formats must stay visible for end users. | Programmatic validation of styling rules before publishing an Excel file to external systems.
// AI Prompts: Generate C# code that applies a 12‑color custom theme with Aspose.Cells and then verifies a between‑value conditional format still highlights the correct cells. | Write a unit‑test in .NET that asserts GetConditionalFormattingResult returns the expected background color after calling Workbook.CustomTheme. | Explain the interaction between Workbook.CustomTheme and conditional formatting in Aspose.Cells, and show how to retrieve the applied style for validation.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeValidation
{
    // The sample creates a workbook, adds a between‑value conditional format to cells A1:A5, applies a 12‑color custom theme with Workbook.CustomTheme, then iterates through the range using GetConditionalFormattingResult to confirm the rule still triggers and the red background is retained before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (rows 1-5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(10 + i * 10); // 10,20,30,40,50
            }

            // Add a conditional formatting that highlights values between 15 and 35
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A5 for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 4,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add the condition and set a style (background red)
            int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "15", "35");
            FormatCondition condition = fcc[conditionIdx];
            condition.Style.BackgroundColor = Color.Red;

            // Apply a custom theme (12 colors required)
            Color[] customColors = new Color[]
            {
                Color.White,          // Background1
                Color.Black,          // Text1
                Color.LightGray,      // Background2
                Color.DarkGray,       // Text2
                Color.Orange,         // Accent1
                Color.Purple,         // Accent2
                Color.Teal,           // Accent3
                Color.Maroon,         // Accent4
                Color.Navy,           // Accent5
                Color.Olive,          // Accent6
                Color.Blue,           // Hyperlink
                Color.Red             // Followed Hyperlink
            };
            workbook.CustomTheme("CustomDemoTheme", customColors);

            // Validate that the conditional formatting still works after applying the theme
            Console.WriteLine("Validation of conditional formatting after theme change:");
            for (int row = 0; row <= 4; row++)
            {
                Cell cell = sheet.Cells[row, 0];
                ConditionalFormattingResult result = cell.GetConditionalFormattingResult();

                // If the cell meets the condition, ConditionalStyle will be non‑null
                bool meetsCondition = result?.ConditionalStyle != null;
                Console.WriteLine($"Cell {cell.Name} value = {cell.Value} => Condition met: {meetsCondition}");

                // Optionally, display the applied background color
                if (meetsCondition)
                {
                    Color appliedColor = result.ConditionalStyle.BackgroundColor;
                    Console.WriteLine($"  Applied background color: {appliedColor}");
                }
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("ThemeValidationResult.xlsx");
        }
    }
}
