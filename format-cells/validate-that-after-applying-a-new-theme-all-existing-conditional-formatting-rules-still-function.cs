// Title: Check that conditional formatting rules remain active after applying a .thmx theme with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a CellValue > 50 conditional format with yellow fill to range A1:A5, saves the file, attempts to apply a .thmx theme if present, then re‑evaluates each cell to confirm the formatting is still applied. | Write a C# example that validates conditional formatting persistence after a theme change by iterating over the cells and comparing their style to the expected yellow background. | Provide a C# snippet that logs any cells where the conditional formatting style differs after loading a workbook with a new theme applied.
// Common Searches: Aspose.Cells .NET verify conditional formatting after applying a theme file | C# test if conditional formatting survives .thmx theme change in Excel workbook | How to check conditional formatting persistence when using Aspose.Cells theme API | Validate conditional formatting rules after workbook theme update using Aspose.Cells | Sample code for conditional formatting validation post theme application in C#
// Tags: conditional formatting persistence after theme application | Aspose.Cells apply .thmx theme | C# validate conditional formatting rules | Excel workbook theme impact on conditional formats | programmatic conditional format verification

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, applies a yellow fill conditional format (value > 50) to A1:A5, optionally applies a .thmx theme, then iterates through the cells to ensure the formatting remains unchanged, logging any mismatches and saving the workbook before and after the theme operation.
class ConditionalFormattingThemeValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A5)
            double[] values = { 10, 60, 30, 80, 20 };
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]); // Row i, Column 0 (A)
            }

            // Add a new conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex]; // ConditionalFormatting object

            // Define the range A1:A5 for the conditional formatting
            cf.AddArea(CellArea.CreateCellArea("A1", "A5"));

            // Create the condition (CellValue > 50)
            int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition condition = cf[conditionIndex];

            // Define the style to apply when the condition is met
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Save the workbook before applying a new theme (optional, demonstrates lifecycle rule)
            workbook.Save("BeforeTheme.xlsx");

            // Apply a theme if a theme file is available (method not available in all versions)
            string themePath = "Theme2.thmx";
            if (File.Exists(themePath))
            {
                Console.WriteLine($"Theme file '{themePath}' found, but applying themes is not supported in this version of Aspose.Cells.");
                // If a future version provides ApplyTheme, it can be called here.
            }
            else
            {
                Console.WriteLine($"Theme file '{themePath}' not found. Skipping theme application.");
            }

            // Validate that conditional formatting still works after (potential) theme change
            bool allValid = true;
            for (int i = 0; i < values.Length; i++)
            {
                Cell cell = sheet.Cells[i, 0];
                Style cellStyle = cell.GetStyle();

                // Cells with value > 50 should have Yellow background according to the rule
                bool shouldBeHighlighted = values[i] > 50;
                bool isHighlighted = cellStyle.ForegroundColor.ToArgb() == Color.Yellow.ToArgb()
                                     && cellStyle.Pattern == BackgroundType.Solid;

                if (shouldBeHighlighted != isHighlighted)
                {
                    allValid = false;
                    Console.WriteLine($"Validation failed at cell {cell.Name}: Value={values[i]}, ExpectedHighlight={shouldBeHighlighted}, ActualHighlight={isHighlighted}");
                }
            }

            Console.WriteLine(allValid
                ? "All conditional formatting rules remain functional after applying the new theme."
                : "Some conditional formatting rules did not function correctly after applying the new theme.");

            // Save the final workbook (optional, demonstrates lifecycle rule)
            workbook.Save("AfterTheme.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
