// Title: Export Excel worksheet with conditional formatting to HTML using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, adds a style rule (value > 50, red background) to cells A1:A10, and saves the sheet as HTML with HtmlSaveOptions. | Show how to configure HtmlSaveOptions so that visual cell styles are retained when converting an Aspose.Cells workbook to an HTML file.
// Common Searches: Aspose.Cells C# export worksheet to HTML while keeping conditional formatting | How to preserve conditional formatting colors when saving Excel as HTML with Aspose | HtmlSaveOptions example for retaining cell style rules in .NET | C# tutorial for red background rule export to HTML using Aspose.Cells
// Tags: conditional formatting to HTML Aspose.Cells | HtmlSaveOptions preserve cell styles | C# export worksheet as HTML Aspose.Cells | red background rule cell value greater than 50 | Aspose.Cells conditional format range example

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, fills A1:A10 with multiples of 10, applies a style rule that colors cells red when the value exceeds 50, and saves the worksheet as an HTML file using HtmlSaveOptions, which retains the visual formatting.
class ConditionalFormattingToHtml
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in range A1:A10
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i * 10); // values: 0,10,20,...,90
            }

            // Define the range to which the formatting will be applied
            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = 10;
            int totalColumns = 1;

            // Add a new conditional formatting rule to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Set the target area (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = firstColumn,
                EndRow = firstRow + totalRows - 1,
                EndColumn = firstColumn + totalColumns - 1
            };
            fcc.AddArea(area);

            // Define the condition: Cell Value > 50
            int conditionIndex = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition condition = fcc[conditionIndex];

            // Define the style to apply when the condition is met (red background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Red;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Prepare HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            string htmlPath = "ConditionalFormattingOutput.html";
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Workbook saved as HTML with conditional formatting at: {htmlPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
