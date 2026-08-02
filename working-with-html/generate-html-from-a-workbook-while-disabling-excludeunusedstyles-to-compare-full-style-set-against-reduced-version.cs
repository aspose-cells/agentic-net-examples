// Title: Export Workbook to HTML with All Styles Preserved (ExcludeUnusedStyles = false) – Aspose.Cells C#
// Description: Shows how to use Aspose.Cells HtmlSaveOptions in C# to generate HTML that retains every cell style by setting ExcludeUnusedStyles to false, producing a full CSS style set for comparison with the reduced output.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExcludeUnusedStyles | C# HTML export | preserve cell styles | full CSS output | Excel to HTML conversion
// Common Searches: Aspose.Cells export to HTML keep all styles | HtmlSaveOptions ExcludeUnusedStyles false example | C# generate HTML from workbook with full CSS | compare Aspose.Cells HTML output with and without unused styles | how to disable style pruning in Aspose.Cells HTML export
// Developer Intent: Generate an HTML file from an Excel workbook while keeping every defined cell style, so the CSS includes both used and unused styles.
// Use Cases: Create two HTML reports (full‑style and minimal) to measure CSS size and rendering performance. | Produce printable HTML documents that must exactly match the original Excel formatting. | Validate that custom fonts, colors, and formatting survive the conversion for quality assurance. | Integrate full‑style HTML output into web applications that rely on a complete stylesheet.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as HTML with ExcludeUnusedStyles set to false, then repeat with true for side‑by‑side comparison. | Explain the impact of the ExcludeUnusedStyles property on the generated CSS and how to inspect the HTML to confirm all style definitions are present. | Provide a unit test in C# that asserts the number of <style> elements differs when ExcludeUnusedStyles is false versus true.

using System;
using Aspose.Cells;
using System.Drawing;

// Shows how to use Aspose.Cells HtmlSaveOptions in C# to generate HTML that retains every cell style by setting ExcludeUnusedStyles to false, producing a full CSS style set for comparison with the reduced output.
class HtmlExportExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data with different styles
        Cell cell1 = workbook.Worksheets[0].Cells["A1"];
        cell1.PutValue("Styled Text");
        Style style1 = workbook.CreateStyle();
        style1.Font.Color = Color.Red;
        style1.Font.IsBold = true;
        cell1.SetStyle(style1);

        Cell cell2 = workbook.Worksheets[0].Cells["A2"];
        cell2.PutValue("Another Style");
        Style style2 = workbook.CreateStyle();
        style2.Font.Name = "Times New Roman";
        style2.Font.Size = 14;
        cell2.SetStyle(style2);

        // Configure HTML save options to keep all styles (disable exclusion of unused styles)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExcludeUnusedStyles = false; // retain full style set in the generated HTML

        // Save the workbook as HTML with the full style set
        workbook.Save("FullStyles.html", saveOptions);
    }
}
