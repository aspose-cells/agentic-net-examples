// Title: Merge V8:W9, Apply Red Conditional Formatting for Negative Values, and Export to HTML – Aspose.Cells C# Example
// Description: A C# snippet that loads an existing workbook, merges the range V8:W9, adds a conditional‑formatting rule that paints cells red when the value is below zero, and saves the result as clean HTML using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | C# conditional formatting negative values | export Excel to HTML Aspose | HtmlSaveOptions MergeAreas | red background conditional format | V8 W9 cell merge | Aspose.Cells .NET example
// Common Searches: how to merge cells V8:W9 with Aspose.Cells C# | apply red conditional formatting for values < 0 in Aspose.Cells | save Excel workbook as HTML preserving formatting | Aspose.Cells conditional formatting HTML output | C# code to merge range and export to HTML
// Developer Intent: Merge the V8:W9 range, highlight negative numbers with a red background, and generate an HTML file from the workbook.
// Use Cases: Create a highlighted header in a financial report where negative figures appear in red, then share the report as a web page. | Convert an Excel dashboard to HTML while keeping merged header cells and loss‑indicating formatting. | Automate batch conversion of spreadsheets to HTML that retains visual cues such as merged cells and conditional styles.
// AI Prompts: Generate C# code using Aspose.Cells to merge V8:W9, add a red background rule for values less than zero, and export the workbook to HTML with MergeAreas enabled. | Explain the effect of HtmlSaveOptions.MergeAreas on the HTML output when conditional formatting is applied in Aspose.Cells. | Provide a step‑by‑step guide to apply multiple conditional‑formatting rules to different ranges before saving the workbook as HTML with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // A C# snippet that loads an existing workbook, merges the range V8:W9, adds a conditional‑formatting rule that paints cells red when the value is below zero, and saves the result as clean HTML using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.html";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // 1. Merge cells V8:W9 (V -> column 21, W -> column 22; rows 8-9 -> indices 7-8)
                // -------------------------------------------------
                cells.Merge(firstRow: 7, firstColumn: 21, totalRows: 2, totalColumns: 2);

                // -------------------------------------------------
                // 2. Apply conditional formatting for negative values on the same range V8:W9
                // -------------------------------------------------
                // Add a new conditional formatting rule to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();

                // Define the target range V8:W9
                CellArea area = new CellArea
                {
                    StartRow = 7,
                    StartColumn = 21,
                    EndRow = 8,
                    EndColumn = 22
                };
                sheet.ConditionalFormattings[cfIndex].AddArea(area);

                // Add a condition: cell value less than 0
                int conditionIndex = sheet.ConditionalFormattings[cfIndex].AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.LessThan,
                    "0",
                    null);

                // Configure the style for the condition (red background)
                FormatCondition condition = sheet.ConditionalFormattings[cfIndex][conditionIndex];
                Style style = condition.Style;
                style.ForegroundColor = Color.Red;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // -------------------------------------------------
                // 3. Save the workbook as HTML
                // -------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Merge conditional formatting areas before saving to keep HTML tidy
                    MergeAreas = true
                };

                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved as HTML to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
