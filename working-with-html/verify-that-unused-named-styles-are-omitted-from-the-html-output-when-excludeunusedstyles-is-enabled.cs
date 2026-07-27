// Title: Exclude Unused Named Styles from HTML with Aspose.Cells HtmlSaveOptions (C#)
// Description: Shows how to create a used and an unused named style in an Aspose.Cells workbook, export the workbook to HTML with HtmlSaveOptions.ExcludeUnusedStyles set to true, and confirm that only the applied style is present in the output.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExcludeUnusedStyles | C# | .NET | Excel to HTML | named style | unused style removal | HTML export | style filtering
// Common Searches: Aspose.Cells exclude unused styles html | HtmlSaveOptions ExcludeUnusedStyles example | remove unused named styles when saving to HTML | verify style omission in Aspose.Cells HTML output | C# export Excel to HTML without redundant CSS
// Developer Intent: Confirm that enabling HtmlSaveOptions.ExcludeUnusedStyles eliminates any named style that is not applied to a cell during HTML conversion.
// Use Cases: Produce clean HTML reports from Excel workbooks by stripping unused CSS definitions. | Automate a quality‑check step that asserts no orphaned style blocks exist in exported HTML. | Reduce page size for web‑based spreadsheet viewers by exporting only necessary styles.
// AI Prompts: Generate a C# unit test using Aspose.Cells that verifies unused named styles are omitted when HtmlSaveOptions.ExcludeUnusedStyles is true. | Provide a step‑by‑step tutorial for exporting an Excel file to HTML with only applied styles, including code to read and validate the HTML content. | Explain the internal mechanism of HtmlSaveOptions.ExcludeUnusedStyles and list scenarios where a style might still appear in the HTML.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a used and an unused named style in an Aspose.Cells workbook, export the workbook to HTML with HtmlSaveOptions.ExcludeUnusedStyles set to true, and confirm that only the applied style is present in the output.
    class VerifyExcludeUnusedStyles
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Create a named style that will be used
            // -------------------------------------------------
            Style usedStyle = workbook.CreateStyle();
            usedStyle.Name = "UsedStyle";
            usedStyle.Font.Color = System.Drawing.Color.Blue;
            usedStyle.Font.IsBold = true;
            // Apply the used style to a cell
            sheet.Cells["A1"].PutValue("Styled Cell");
            sheet.Cells["A1"].SetStyle(usedStyle);

            // -------------------------------------------------
            // Create a named style that will NOT be used
            // -------------------------------------------------
            Style unusedStyle = workbook.CreateStyle();
            unusedStyle.Name = "UnusedStyle";
            unusedStyle.Font.Color = System.Drawing.Color.Red;
            unusedStyle.Font.IsItalic = true;
            // NOTE: Do NOT apply this style to any cell

            // -------------------------------------------------
            // Configure HTML save options to exclude unused styles
            // -------------------------------------------------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExcludeUnusedStyles = true; // Ensure unused styles are omitted

            // Define output HTML file path
            string htmlPath = "output.html";

            // Save the workbook as HTML with the configured options
            workbook.Save(htmlPath, saveOptions);

            // -------------------------------------------------
            // Verify that the unused style is not present in the HTML
            // -------------------------------------------------
            string htmlContent = File.ReadAllText(htmlPath);

            bool containsUsedStyle = htmlContent.Contains("UsedStyle");
            bool containsUnusedStyle = htmlContent.Contains("UnusedStyle");

            Console.WriteLine($"HTML contains used style (should be true): {containsUsedStyle}");
            Console.WriteLine($"HTML contains unused style (should be false): {containsUnusedStyle}");

            // Simple assertion output
            if (containsUsedStyle && !containsUnusedStyle)
            {
                Console.WriteLine("Verification succeeded: Unused named styles are omitted.");
            }
            else
            {
                Console.WriteLine("Verification failed: Unused named styles are present.");
            }
        }
    }
}
