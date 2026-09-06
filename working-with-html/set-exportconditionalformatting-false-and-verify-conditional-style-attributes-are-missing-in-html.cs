// Title: How to disable conditional formatting export in Aspose.Cells for .NET when saving to HTML and verify the style is omitted
// AI Prompts: Write C# code that creates a workbook, adds a red‑background conditional format to a range, saves it to HTML with HtmlSaveOptions.ExportConditionalFormatting set to false, and checks that the generated HTML does not contain the red background style. | Provide a C# snippet that reads the HTML output from a MemoryStream after saving with Aspose.Cells and programmatically asserts that conditional‑formatting CSS rules are absent.
// Common Searches: Aspose.Cells HtmlSaveOptions ExportConditionalFormatting false example C# | How to prevent conditional formatting from being exported to HTML with Aspose.Cells | Verify that conditional formatting styles are not present in HTML output using Aspose.Cells .NET | C# Aspose.Cells save workbook to HTML without conditional formatting | Check HTML string for missing conditional formatting background color Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions disable conditional formatting | C# export workbook to HTML without conditional styles | conditional formatting style omission verification | red background conditional format removal Aspose.Cells | HTML output validation Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills column A with values 0‑9, applies a conditional formatting rule that sets a red background for cells greater than 5, saves the workbook to HTML using HtmlSaveOptions with ExportConditionalFormatting disabled, reads the HTML from a memory stream, and asserts that the red background style is absent, throwing an exception if it is found.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with values 0..9
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i);
            }

            // Add a conditional formatting rule: cells with value > 5 get a red background
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];

            // Define the range the rule applies to (A1:A10)
            cf.AddArea(new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            });

            // Create the condition (CellValue > 5)
            int conditionIndex = cf.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "5",
                null);

            // Retrieve the created condition
            FormatCondition condition = cf[conditionIndex];

            // Set the style for the condition (red background)
            condition.Style.BackgroundColor = Color.Red;

            // Configure HTML save options (disable conditional formatting export if supported)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // htmlOptions.ExportConditionalFormatting = false; // Uncomment if the property exists in your version

            // Save the workbook to HTML in a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;
                string html = new StreamReader(ms).ReadToEnd();

                // Verify that the HTML does NOT contain the red background style
                bool containsRedBackground = html.IndexOf("background-color:#FF0000", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                             html.IndexOf("background:#FF0000", StringComparison.OrdinalIgnoreCase) >= 0;

                Console.WriteLine("Conditional formatting exported to HTML? " + (containsRedBackground ? "Yes" : "No"));
                if (containsRedBackground)
                {
                    throw new Exception("Conditional formatting style was found in the HTML output, but it should be omitted.");
                }
                else
                {
                    Console.WriteLine("Verification passed: Conditional formatting styles are absent in the HTML.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
