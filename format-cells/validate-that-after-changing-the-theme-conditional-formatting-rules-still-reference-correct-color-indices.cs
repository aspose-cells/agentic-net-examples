// Title: Aspose.Cells .NET – Verify Conditional Formatting Keeps Accent1 Theme Color After Workbook Theme Change
// Description: C# example that creates a workbook, applies a conditional formatting rule to A1:A5 with a font styled by ThemeColorType.Accent1, changes the workbook's Accent1 theme color using SetThemeColor, and validates that the rule still references the Accent1 theme type and reports its tint value.
// Keywords: Aspose.Cells | C# | conditional formatting | theme color | Accent1 | SetThemeColor | ThemeColorType | font style validation | Excel theme change | .NET workbook
// Common Searches: Aspose.Cells verify conditional formatting theme after SetThemeColor | C# check if conditional formatting keeps Accent1 after theme change | how to validate theme color reference in Aspose.Cells conditional formatting | Aspose.Cells .NET conditional formatting theme color consistency
// Developer Intent: Ensure that conditional formatting rules continue to reference the original ThemeColorType (Accent1) after the workbook's theme colors are modified.
// Use Cases: Programmatically confirm theme color integrity of existing conditional formatting after applying a new workbook theme. | Automated testing of theme‑dependent styling in Excel files generated with Aspose.Cells. | Detect and correct mismatched theme references when dynamically updating workbook themes in .NET applications.
// AI Prompts: Generate C# code with Aspose.Cells that iterates all conditional formatting rules and asserts each Font.ThemeColor.ColorType remains Accent1 after calling SetThemeColor. | Explain the impact of Workbook.SetThemeColor on ThemeColor objects used in conditional formatting and describe how to validate them. | Create an MSTest unit test that adds a conditional formatting rule using ThemeColorType.Accent1, changes the theme color, and verifies the ThemeColorType is unchanged and the tint value is as expected.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeConditionalFormattingValidation
{
    // C# example that creates a workbook, applies a conditional formatting rule to A1:A5 with a font styled by ThemeColorType.Accent1, changes the workbook's Accent1 theme color using SetThemeColor, and validates that the rule still references the Accent1 theme type and reports its tint value.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 5); // values: 0,5,10,15,20
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A5 for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 0,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Add a condition: Cell value greater than 10
            int conditionIdx = cfCollection.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
            FormatCondition condition = cfCollection[conditionIdx];

            // Set the condition style to use a theme color (Accent1) for the font
            condition.Style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);

            // OPTIONAL: Save the workbook before changing the theme (demonstration purpose)
            // workbook.Save("BeforeThemeChange.xlsx");

            // Change the theme color for Accent1 to a different color (e.g., Green)
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Green);

            // Validate that the conditional formatting still references the correct theme color type
            ThemeColor themeColorAfterChange = condition.Style.Font.ThemeColor;

            bool isCorrectThemeType = themeColorAfterChange != null && themeColorAfterChange.ColorType == ThemeColorType.Accent1;

            Console.WriteLine("Conditional formatting font still uses ThemeColorType.Accent1: " + isCorrectThemeType);
            Console.WriteLine("Tint value after change: " + themeColorAfterChange.Tint);

            // OPTIONAL: Save the workbook after theme change
            // workbook.Save("AfterThemeChange.xlsx");
        }
    }
}
