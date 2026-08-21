// Title: C# Aspose.Cells – Conditional Formatting to Highlight Cells with QuotePrefix = False
// Description: This example creates a workbook, inserts values with and without leading apostrophes, defines a conditional‑formatting range (A1:A4), and applies an expression‑based rule using CELL("prefix",A1)="" to detect cells where the QuotePrefix property is false. Matching cells receive a LightYellow solid background and the workbook is saved as QuotePrefixConditionalFormatting.xlsx.
// Keywords: Aspose.Cells C# conditional formatting | QuotePrefix false detection | Excel CELL prefix formula | highlight cells without leading apostrophe | set cell background color Aspose.Cells | expression based conditional formatting
// Common Searches: Aspose.Cells conditional formatting QuotePrefix false | C# detect leading apostrophe in Excel cells | apply background color when QuotePrefix is false | CELL function conditional formatting Aspose | how to highlight cells without a quote prefix using Aspose.Cells
// Developer Intent: Apply a conditional‑formatting rule that colors cells whose QuotePrefix property is false.
// Use Cases: Flag entries entered without a leading apostrophe for data‑entry validation. | Visually separate raw text from quoted strings in reports. | Enforce consistent formatting by automatically shading cells lacking a QuotePrefix.
// AI Prompts: Generate C# code with Aspose.Cells that colors cells where QuotePrefix is false using a custom formula. | Suggest an alternative Excel formula for detecting QuotePrefix false in Aspose.Cells conditional formatting. | Explain how to modify the background color, pattern, or range of the conditional formatting rule in the provided example.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingDemo
{
    // This example creates a workbook, inserts values with and without leading apostrophes, defines a conditional‑formatting range (A1:A4), and applies an expression‑based rule using CELL("prefix",A1)="" to detect cells where the QuotePrefix property is false. Matching cells receive a LightYellow solid background and the workbook is saved as QuotePrefixConditionalFormatting.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // A1 – normal text (QuotePrefix = false)
            cells["A1"].PutValue("Hello");
            // A2 – text with leading apostrophe (QuotePrefix = true)
            cells["A2"].PutValue("'World");
            // A3 – numeric value (QuotePrefix = false)
            cells["A3"].PutValue(123);
            // A4 – numeric value with leading apostrophe (QuotePrefix = true)
            cells["A4"].PutValue("'456");

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will be applied (A1:A4)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 3,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add an expression‑based condition that evaluates to TRUE when QuotePrefix is FALSE
            // Excel's CELL("prefix",A1) returns a single quote (') if QuotePrefix is true, otherwise empty.
            // The formula checks for an empty result, meaning QuotePrefix is false.
            string formula = "CELL(\"prefix\",A1)=\"\"";
            int conditionIndex = fcs.AddCondition(FormatConditionType.Expression, OperatorType.None, formula, null);
            FormatCondition condition = fcs[conditionIndex];

            // Set a distinct background color for cells that meet the condition
            condition.Style.BackgroundColor = Color.LightYellow;
            condition.Style.Pattern = BackgroundType.Solid;

            // Save the workbook
            workbook.Save("QuotePrefixConditionalFormatting.xlsx");
        }
    }
}
