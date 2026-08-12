// Title: Apply Conditional Formatting and Export to HTML with a Custom TableCssId using Aspose.Cells for .NET
// Description: Creates a workbook, fills cells A1:B3, adds a conditional formatting rule that colors values > 50 red, sets HtmlSaveOptions.TableCssId to "customTable", saves the workbook as HTML, and verifies that the generated HTML contains the specified TableCssId.
// Keywords: Aspose.Cells | C# conditional formatting | HTML export | TableCssId | save workbook as HTML | verify HTML output | conditional formatting preservation | Aspose.Cells HtmlSaveOptions
// Common Searches: Aspose.Cells conditional formatting export to HTML | How to set TableCssId in HtmlSaveOptions | Verify TableCssId in generated HTML with Aspose.Cells | Preserve conditional formatting colors in HTML output | C# Aspose.Cells HTML export custom table id
// Developer Intent: Add a conditional formatting rule to a range and ensure the formatting and custom TableCssId are retained when the workbook is saved as HTML.
// Use Cases: Generate HTML reports where values above a threshold are highlighted and the table can be styled via a specific CSS selector. | Automate unit tests that confirm the exported HTML includes the configured TableCssId for downstream processing. | Create multiple HTML files from different worksheets, each using a unique TableCssId to apply distinct CSS themes.
// AI Prompts: Show how to add multiple conditional formatting rules and export each worksheet to HTML with different TableCssId values using Aspose.Cells. | Provide C# code that parses the saved HTML and asserts that cells with values > 50 have the red background style applied. | Explain how to customize the CSS classes generated for conditional formatting while keeping the custom TableCssId in the HTML output.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// Creates a workbook, fills cells A1:B3, adds a conditional formatting rule that colors values > 50 red, sets HtmlSaveOptions.TableCssId to "customTable", saves the workbook as HTML, and verifies that the generated HTML contains the specified TableCssId.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the range A1:B3
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(60);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["B1"].PutValue(70);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["B3"].PutValue(90);

        // Add conditional formatting: highlight cells with value > 50 in red background
        int cfIndex = sheet.ConditionalFormattings.Add();                     // Create a new ConditionalFormatting object
        var cf = sheet.ConditionalFormattings[cfIndex];                     // Retrieve the collection
        var area = new CellArea                                            // Define the target range A1:B3
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 2,
            EndColumn = 1
        };
        cf.AddArea(area);                                                  // Apply the range to the conditional formatting
        int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
        var condition = cf[conditionIndex];                                 // Get the created condition
        condition.Style.BackgroundColor = Color.Red;                        // Set the style for the condition

        // Configure HTML save options with a custom TableCssId
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "customTable";

        // Save the workbook as HTML using the configured options
        string htmlPath = "ConditionalFormatting.html";
        workbook.Save(htmlPath, saveOptions);
        Console.WriteLine($"Workbook saved to HTML: {htmlPath}");

        // Verify that the generated HTML contains the specified TableCssId
        string htmlContent = File.ReadAllText(htmlPath);
        bool containsTableCssId = htmlContent.Contains(saveOptions.TableCssId);
        Console.WriteLine($"Verification - TableCssId present in HTML: {containsTableCssId}");
    }
}
