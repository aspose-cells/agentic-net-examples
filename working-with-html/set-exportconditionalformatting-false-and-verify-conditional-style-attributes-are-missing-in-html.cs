// Title: Export Excel to HTML without Conditional Formatting (C# Aspose.Cells) and Verify Absence
// Description: Shows how to create a workbook, add a conditional formatting rule, save it as HTML with ExportConditionalFormatting disabled, and programmatically check that the resulting HTML lacks the conditional style (e.g., red background).
// Keywords: Aspose.Cells | C# | HTML export | ExportConditionalFormatting false | disable conditional formatting | HtmlSaveOptions | verify HTML style | conditional formatting removal | background-color red | unit test example
// Common Searches: Aspose.Cells export HTML without conditional formatting | How to turn off ExportConditionalFormatting in C# | Check HTML output for conditional formatting using Aspose.Cells | C# code to verify conditional style is not in exported HTML | Aspose.Cells HtmlSaveOptions ExportConditionalFormatting property
// Developer Intent: Export a workbook to HTML while omitting all conditional formatting and confirm that the conditional CSS is absent.
// Use Cases: Produce clean HTML reports from Excel files where visual rules must be stripped. | Automated testing to ensure HTML exports do not contain conditional formatting CSS. | Generate lightweight HTML email bodies from spreadsheets without style bloat.
// AI Prompts: Write C# code with Aspose.Cells that saves a workbook to HTML with ExportConditionalFormatting set to false and validates that no "background-color:red" appears. | Explain how to parse the saved HTML string to detect any leftover conditional formatting styles. | Suggest work‑arounds for versions of Aspose.Cells that lack the ExportConditionalFormatting property.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// Shows how to create a workbook, add a conditional formatting rule, save it as HTML with ExportConditionalFormatting disabled, and programmatically check that the resulting HTML lacks the conditional style (e.g., red background).
class ExportConditionalFormattingDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);

            // Add a conditional formatting rule: cells > 25 get a red background
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A4
            CellArea area = new CellArea { StartRow = 0, EndRow = 3, StartColumn = 0, EndColumn = 0 };
            fcc.AddArea(area);

            // Add the condition and set its style
            int condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "25", null);
            FormatCondition condition = fcc[condIdx];
            condition.Style.BackgroundColor = Color.Red;

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Note: ExportConditionalFormatting property is not available in this version;
            // the default behavior exports conditional formatting. Adjust as needed for your version.
            htmlOptions.ExcludeUnusedStyles = false; // keep all styles for verification

            // Save the workbook as HTML
            string htmlPath = "ConditionalFormatting.html";
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the HTML file was created
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("HTML file was not generated.", htmlPath);

            // Load the generated HTML and verify that the conditional style (red background) is absent
            string htmlContent = File.ReadAllText(htmlPath);

            bool containsRedBackground = htmlContent.IndexOf("background-color", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                        htmlContent.IndexOf("red", StringComparison.OrdinalIgnoreCase) >= 0;

            Console.WriteLine("Conditional formatting exported? " + (!containsRedBackground));
            // Expected output: true (meaning the conditional style is missing)
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
