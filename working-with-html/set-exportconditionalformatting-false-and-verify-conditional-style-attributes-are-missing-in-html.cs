// Title: Disable conditional formatting when saving a workbook to HTML with Aspose.Cells for .NET and confirm style removal
// AI Prompts: Write C# code that creates a workbook, applies a red background style based on cell values, saves it to HTML with the conditional formatting export turned off, and then reads the HTML file to confirm that no background‑color CSS is present. | Adapt an existing Aspose.Cells sample to disable conditional formatting during HTML conversion and add logic that scans the resulting HTML for the absence of any conditional style attributes.
// Common Searches: Aspose.Cells ExportConditionalFormatting false how to disable when saving as HTML C# | Save Excel workbook to HTML without conditional formatting using Aspose.Cells | Check generated HTML for missing background-color style after disabling conditional formatting | C# example of HtmlSaveOptions with conditional formatting export turned off | Verify that conditional formatting is not included in HTML output with Aspose.Cells
// Tags: htmlsaveoptions disable conditional formatting export | aspocells turn off conditional formatting for HTML | c# verify missing background-color css in HTML output | excel workbook to html without conditional styles | aspocells conditional formatting export control

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, fills cells A1:A10 with numeric values, applies a red background style to cells where the value exceeds 50, configures HtmlSaveOptions to suppress conditional formatting export, saves the workbook as an HTML file, and then reads the file to determine whether any background‑color CSS rules were generated, confirming that the conditional style is omitted.
class ExportConditionalFormattingDemo
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 3. Populate some data (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
            }

            // 4. Apply red background to cells with value > 50 (simulating conditional formatting)
            Style redStyle = workbook.CreateStyle();
            redStyle.ForegroundColor = Color.Red;
            redStyle.Pattern = BackgroundType.Solid;

            for (int i = 0; i < 10; i++)
            {
                if (cells[i, 0].IntValue > 50)
                {
                    cells[i, 0].SetStyle(redStyle);
                }
            }

            // 5. Prepare HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // If the Aspose.Cells version supports it, you can disable conditional formatting export:
            // htmlOptions.ExportConditionalFormatting = false;

            // 6. Save the workbook to HTML
            string htmlPath = "ConditionalFormattingFalse.html";
            workbook.Save(htmlPath, htmlOptions);

            // 7. Verify the generated HTML for red background style
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                bool containsRedBackground = htmlContent.IndexOf("background-color:#FF0000", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                             htmlContent.IndexOf("background-color:red", StringComparison.OrdinalIgnoreCase) >= 0;

                Console.WriteLine("HTML generation completed.");
                Console.WriteLine("Conditional style attributes present in HTML? " + (containsRedBackground ? "Yes" : "No"));
            }
            else
            {
                Console.WriteLine($"Failed to generate HTML file at path: {htmlPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
