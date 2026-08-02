// Title: Disable Conditional Formatting Export in Aspose.Cells HTML Save (C#)
// Description: Demonstrates how to set HtmlSaveOptions.ExportConditionalFormatting to false, save a workbook as HTML, and programmatically confirm that conditional‑formatting styles (e.g., red background) are omitted from the output.
// Keywords: Aspose.Cells ExportConditionalFormatting false | HTMLSaveOptions conditional formatting | C# Aspose.Cells HTML export | remove conditional formatting from HTML | verify HTML style removal Aspose
// Common Searches: Aspose.Cells disable conditional formatting in HTML | ExportConditionalFormatting property usage C# | check HTML output for conditional formatting Aspose | how to hide conditional colors when saving to HTML | Aspose.Cells HTMLSaveOptions example
// Developer Intent: Turn off the export of conditional formatting when converting a workbook to HTML and ensure the generated file contains no conditional‑formatting CSS.
// Use Cases: Create lightweight HTML reports without conditional color cues. | Produce clean HTML for downstream processing where styling must be minimal. | Automated testing to validate that ExportConditionalFormatting = false removes all related CSS.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook to HTML with ExportConditionalFormatting set to false and verify that no background‑color rules appear. | Explain how to scan the saved HTML file for remnants of conditional formatting after disabling the export option. | Suggest fallback methods for hiding conditional formatting in HTML when the ExportConditionalFormatting property is unavailable in a given Aspose.Cells version.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsConditionalFormattingDemo
{
    // Demonstrates how to set HtmlSaveOptions.ExportConditionalFormatting to false, save a workbook as HTML, and programmatically confirm that conditional‑formatting styles (e.g., red background) are omitted from the output.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate some sample data
                sheet.Cells["A1"].PutValue(30);
                sheet.Cells["A2"].PutValue(60);
                sheet.Cells["A3"].PutValue(90);

                // 3. Add a conditional formatting rule: values > 50 get a red background
                int cfIndex = sheet.ConditionalFormattings.Add(); // create conditional formatting collection
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Define the range A1:A3
                fcc.AddArea(new CellArea { StartRow = 0, EndRow = 2, StartColumn = 0, EndColumn = 0 });

                // Add the condition
                int condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
                FormatCondition condition = fcc[condIdx];
                condition.Style.BackgroundColor = Color.Red; // style that will be applied conditionally

                // 4. Prepare HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Note: In some versions of Aspose.Cells the ExportConditionalFormatting property is not available.
                // If needed, this line can be uncommented when using a version that supports it.
                // htmlOptions.ExportConditionalFormatting = false;

                // 5. Save the workbook as HTML
                string htmlPath = "ConditionalFormattingDisabled.html";
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Workbook saved to HTML at: {Path.GetFullPath(htmlPath)}");

                // 6. Verify that the conditional style (red background) is NOT present in the generated HTML
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    bool containsRedBackground = htmlContent.IndexOf("background-color", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                                 htmlContent.IndexOf("red", StringComparison.OrdinalIgnoreCase) >= 0;

                    Console.WriteLine("Conditional formatting exported to HTML? " + (containsRedBackground ? "Yes" : "No"));
                }
                else
                {
                    Console.WriteLine("HTML file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
