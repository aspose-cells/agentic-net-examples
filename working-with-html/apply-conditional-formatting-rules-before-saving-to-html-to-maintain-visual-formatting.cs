// Title: Apply red background conditional formatting to cells > 50 and save worksheet as HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add conditional formatting rule for values greater than 50 before exporting to HTML | How to preserve cell background colors when converting Excel to HTML using Aspose.Cells | C# conditional formatting range A1:A10 red fill and save as HTML with Aspose.Cells | Apply cell value based conditional formatting and generate HTML output in Aspose.Cells .NET | Set solid red background for cells > 50 and export workbook to HTML using Aspose.Cells API
// Tags: FormatConditionCollection conditional formatting Aspose.Cells | red background style for cell value greater than 50 C# | save workbook as HTML with conditional formats Aspose.Cells | apply conditional formatting to range A1:A10 Aspose.Cells | cell value based conditional formatting Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example loads an Excel workbook, defines the range A1:A10, creates a conditional formatting rule that applies a solid red background to cells with values greater than 50, and then saves the worksheet as an HTML file, ensuring the formatting is retained in the HTML output.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range to which the conditional formatting will be applied (e.g., A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            StartColumn = 0,   // Column A (zero‑based)
            EndRow = 9,        // Row 10
            EndColumn = 0      // Column A
        };

        // Add a new ConditionalFormatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

        // Associate the defined range with the collection
        cfCollection.AddArea(area);

        // Create a condition: cells with value > 50 will have a red background
        int conditionIndex = cfCollection.AddCondition(
            FormatConditionType.CellValue,          // Type of condition
            OperatorType.GreaterThan,               // Operator
            "50",                                   // First operand (value to compare)
            null);                                  // Second operand (not used)

        FormatCondition condition = cfCollection[conditionIndex];

        // Define the style for the condition
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.Red;
        style.Pattern = BackgroundType.Solid;
        condition.Style = style;

        // Save the workbook as HTML (replace with desired output path)
        workbook.Save("output.html", SaveFormat.Html);
    }
}
