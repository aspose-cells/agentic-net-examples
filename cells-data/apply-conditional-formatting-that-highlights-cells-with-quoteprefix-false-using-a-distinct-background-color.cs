// Title: Apply conditional formatting in Aspose.Cells (.NET) to highlight cells where QuotePrefix is false (no leading single quote)
// AI Prompts: Generate C# code that creates an expression‑based conditional format in Aspose.Cells to color cells whose first character is not a single quote. | Show how to set a light‑yellow background style for cells with QuotePrefix false using Aspose.Cells conditional formatting. | Write the formula and style configuration needed to highlight non‑prefixed cells in an Aspose.Cells worksheet.
// Common Searches: Aspose.Cells C# conditional formatting based on QuotePrefix property | How to highlight cells without a leading apostrophe using Aspose.Cells .NET | Expression condition LEFT(A1,1)<>"'" in Aspose.Cells conditional formatting example | Apply background color to cells where QuotePrefix is false in a workbook with Aspose.Cells
// Tags: formula driven conditional formatting Aspose.Cells C# | cells without QuotePrefix formatting Aspose.Cells | conditional format for non‑prefixed cells .NET | background color style Aspose.Cells conditional format | LEFT function rule Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingQuotePrefix
{
    // The example creates a workbook, adds sample data (A1 with a leading single quote and A2‑A4 without), defines an expression‑based conditional formatting rule over range A1:A4 that uses the formula LEFT(A1,1)<>"'", applies a light‑yellow background style to cells where the condition is true (QuotePrefix false), and saves the file as QuotePrefixConditionalFormatting.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            // Cell with QuotePrefix = true (starts with a single quote)
            worksheet.Cells["A1"].PutValue("'Prefixed text");
            // Cells with QuotePrefix = false
            worksheet.Cells["A2"].PutValue("Normal text 1");
            worksheet.Cells["A3"].PutValue("Normal text 2");
            worksheet.Cells["A4"].PutValue("Normal text 3");

            // Add a conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied (A1:A4)
            CellArea area = new CellArea
            {
                StartRow = 0,   // Row 0 (A1)
                EndRow = 3,     // Row 3 (A4)
                StartColumn = 0,
                EndColumn = 0
            };
            conditions.AddArea(area);

            // Add an expression‑type condition
            // The formula checks whether the first character of the cell is NOT a single quote
            // LEFT(A1,1)<> "'"  -> true when QuotePrefix is false
            int conditionIndex = conditions.AddCondition(FormatConditionType.Expression);
            FormatCondition condition = conditions[conditionIndex];
            condition.Formula1 = "LEFT(A1,1)<>\"'\"";

            // Define the style to apply when the condition is true (e.g., light yellow background)
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.Pattern = BackgroundType.Solid;
            highlightStyle.ForegroundColor = Color.LightYellow;
            condition.Style = highlightStyle;

            // Save the workbook
            workbook.Save("QuotePrefixConditionalFormatting.xlsx");
        }
    }
}
