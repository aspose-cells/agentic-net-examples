// Title: Apply the workbook's Accent2 theme color as a conditional formatting fill based on a formula using Aspose.Cells for .NET
// AI Prompts: Create a conditional formatting rule for the range A1:A100 that fills cells with the workbook's Accent2 theme color when the formula =A1>10 evaluates to true in C# with Aspose.Cells. | Retrieve the Accent2 color from the workbook's theme and build a solid fill style to assign to a conditional formatting condition using Aspose.Cells. | Save the workbook to a new file after adding the theme‑based conditional formatting rule with Aspose.Cells.
// Common Searches: aspocells how to use theme accent colors in conditional formatting | c# apply workbook theme Accent2 as fill color for conditional format | retrieve Excel theme color Accent2 with Aspose.Cells .NET | conditional formatting expression A1>10 using theme color in Aspose.Cells | set solid background from workbook theme in Aspose.Cells C#
// Tags: conditional formatting theme accent color Aspose.Cells | retrieve workbook theme color C# Aspose.Cells | solid fill style for expression condition Aspose.Cells | Excel theme Accent2 fill Aspose.Cells .NET | apply theme‑based conditional format Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel file, adds a conditional formatting rule to cells A1:A100 that uses the workbook's Accent2 theme color as a solid fill when the expression A1>10 is true, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for conditional formatting (A1:A100)
            CellArea area = new CellArea
            {
                StartRow = 0,      // Row index is zero‑based
                StartColumn = 0,   // Column A
                EndRow = 99,       // Row 100 (zero‑based index 99)
                EndColumn = 0      // Column A
            };

            // Add a new ConditionalFormatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();

            // Add the target area to the conditional formatting
            sheet.ConditionalFormattings[cfIndex].AddArea(area);

            // Add a condition based on a formula (e.g., =A1>10)
            int conditionIndex = sheet.ConditionalFormattings[cfIndex].AddCondition(FormatConditionType.Expression);
            FormatCondition condition = sheet.ConditionalFormattings[cfIndex][conditionIndex];
            condition.Formula1 = "A1>10";

            // Retrieve the theme's Accent2 color (compatible with all versions)
            Color accent2Color = workbook.GetThemeColor(ThemeColorType.Accent2);

            // Create a style that uses the Accent2 color as a solid fill
            Style accentStyle = workbook.CreateStyle();
            accentStyle.ForegroundColor = accent2Color;
            accentStyle.Pattern = BackgroundType.Solid;

            // Assign the style to the condition
            condition.Style = accentStyle;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
