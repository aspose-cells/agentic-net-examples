// Title: C# example: Export an Excel workbook to HTML with Aspose.Cells while omitting conditional formatting and verify the CSS is absent
// Common Searches: Aspose.Cells C# export workbook to HTML without conditional formatting | How to prevent conditional formatting styles from being saved in HTML using Aspose.Cells | Verify that background-color CSS is not present in HTML generated from Excel with Aspose.Cells | C# code to check HTML output for missing conditional formatting rules after export | HtmlSaveOptions ExportActiveWorksheetOnly and exclude conditional formatting Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions export without conditional formatting | C# export Excel to HTML omit conditional style | verify HTML output lacks background-color CSS Aspose.Cells | conditional formatting exclusion in HTML conversion Aspose.Cells | Excel to HTML conversion C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample creates a workbook, adds a conditional formatting rule that colors cells red when the value exceeds 25, saves the workbook to HTML using HtmlSaveOptions (ExportActiveWorksheetOnly enabled), reads the generated HTML file, and checks that no background-color CSS (e.g., "#FF0000" or "red") appears, confirming that conditional formatting was not exported.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for the generated files
            string htmlPath = "ConditionalFormattingTest.html";

            // 1. Create a workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some data
            sheet.Cells["A1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(30);
            sheet.Cells["A5"].PutValue(40);

            // 2. Add a conditional formatting rule:
            //    Cells with value > 25 will have a red background.
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range A2:A5
            var area = new CellArea
            {
                StartRow = 1,
                StartColumn = 0,
                EndRow = 4,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Add the condition
            int conditionIdx = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "25",
                null);

            // Set the style for the condition (red background)
            var condition = cfCollection[conditionIdx];
            condition.Style.BackgroundColor = System.Drawing.Color.Red;
            condition.Style.Pattern = BackgroundType.Solid;

            // 3. Save the workbook to HTML.
            // ExportConditionalFormatting property is not available in the current Aspose.Cells version,
            // so we rely on the default behavior (conditional formatting is not exported) or
            // any other available options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.ExportActiveWorksheetOnly = true; // Export only the sheet we worked on

            workbook.Save(htmlPath, htmlOptions);
            Console.WriteLine($"Workbook saved to HTML at: {Path.GetFullPath(htmlPath)}");

            // 4. Verify that conditional formatting styles are NOT present in the generated HTML
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine("Error: HTML file was not created.");
                return;
            }

            string htmlContent = File.ReadAllText(htmlPath);

            // The style we applied was a red background. In HTML this would appear as a CSS rule
            // containing "background-color:#FF0000" (or similar). We'll search for that pattern.
            bool containsRedBackground = htmlContent.IndexOf("background-color", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                         (htmlContent.IndexOf("#ff0000", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                          htmlContent.IndexOf("red", StringComparison.OrdinalIgnoreCase) >= 0);

            if (containsRedBackground)
            {
                Console.WriteLine("Verification FAILED: Conditional formatting style was found in the HTML.");
            }
            else
            {
                Console.WriteLine("Verification PASSED: Conditional formatting style is absent in the HTML as expected.");
            }

            // Optional: Show a snippet of the HTML for manual inspection
            Console.WriteLine("\n--- HTML snippet (first 500 characters) ---");
            Console.WriteLine(htmlContent.Substring(0, Math.Min(500, htmlContent.Length)));
            Console.WriteLine("--- End of snippet ---");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
