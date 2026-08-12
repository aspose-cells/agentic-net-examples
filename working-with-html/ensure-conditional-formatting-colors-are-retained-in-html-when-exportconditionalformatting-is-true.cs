// Title: Preserve Conditional Formatting Colors in HTML Export Using Aspose.Cells for .NET
// Description: Shows how to add a red‑background conditional format (value > 15) to a workbook and set HtmlSaveOptions (ExcludeUnusedStyles = false) so the formatting is kept when the file is saved as HTML.
// Keywords: Aspose.Cells | C# | .NET | HTML export | conditional formatting | preserve colors | HtmlSaveOptions | ExcludeUnusedStyles | ExportConditionalFormatting | Excel to HTML | web report
// Common Searches: Aspose.Cells keep conditional formatting colors in HTML | HTML export loses conditional formatting Aspose.Cells | ExportConditionalFormatting true not working | How to retain conditional formatting when saving Excel as HTML | HtmlSaveOptions ExcludeUnusedStyles false example
// Developer Intent: The developer needs the conditional formatting (e.g., red background) to appear in the generated HTML file.
// Use Cases: Create web‑based financial dashboards that highlight out‑of‑range values. | Generate email‑ready HTML reports that preserve Excel visual cues. | Build interactive data tables for intranet portals with Excel‑style conditional highlights.
// AI Prompts: Provide C# code that ensures conditional formatting colors are exported to HTML with Aspose.Cells. | Explain the role of ExcludeUnusedStyles and ExportConditionalFormatting when saving a workbook as HTML. | Give troubleshooting steps for missing conditional formatting colors in Aspose.Cells HTML output.

using System;
using System.Drawing;
using Aspose.Cells;

// Shows how to add a red‑background conditional format (value > 15) to a workbook and set HtmlSaveOptions (ExcludeUnusedStyles = false) so the formatting is kept when the file is saved as HTML.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a conditional formatting rule: values greater than 15 will have a red background
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range the rule applies to (A1:A3)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Create the condition and set its style
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "15", null);
            FormatCondition condition = cfCollection[conditionIndex];
            condition.Style.BackgroundColor = Color.Red; // conditional color

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // ExportConditionalFormatting property is not required; conditional formatting is exported by default
                ExcludeUnusedStyles = false // keep all styles so conditional styles are not stripped
            };

            // Save the workbook as HTML
            string outputPath = "ConditionalFormatting.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
