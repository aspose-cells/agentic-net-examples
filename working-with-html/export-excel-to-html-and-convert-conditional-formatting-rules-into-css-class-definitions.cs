// Title: Export Excel to HTML with Conditional Formatting as CSS Classes using Aspose.Cells for .NET
// Description: Loads an Excel workbook, configures HtmlSaveOptions to export each worksheet's CSS separately, includes all data and styles, disables removal of unused styles, and saves the result as an HTML file so that every conditional‑formatting rule is emitted as a CSS class definition.
// Keywords: Aspose.Cells HTML export | conditional formatting to CSS | ExportWorksheetCSSSeparately | ExcludeUnusedStyles false | HtmlExportDataOptions All | C# Excel to HTML conversion | generate CSS from Excel
// Common Searches: Aspose.Cells export Excel to HTML with CSS classes | keep conditional formatting when saving Excel as HTML | how to generate separate CSS for Excel worksheets | prevent unused style removal Aspose.Cells HTML | C# convert Excel conditional formatting to CSS
// Developer Intent: Create an HTML representation of an Excel workbook where all conditional‑formatting rules are preserved as reusable CSS class definitions.
// Use Cases: Web dashboards that need the same visual cues as the original Excel report. | Archiving Excel workbooks as static HTML pages while retaining conditional formatting. | Building responsive web pages that load worksheet‑specific CSS for faster styling overrides.
// AI Prompts: Generate C# code with Aspose.Cells to export an Excel file to HTML, converting conditional formatting into separate CSS class files. | Explain the impact of ExportWorksheetCSSSeparately, ExportDataOptions, and ExcludeUnusedStyles on conditional formatting during HTML export. | Show how to modify the sample to embed the generated CSS inline instead of creating external CSS files.

using System;
using Aspose.Cells;

// Loads an Excel workbook, configures HtmlSaveOptions to export each worksheet's CSS separately, includes all data and styles, disables removal of unused styles, and saves the result as an HTML file so that every conditional‑formatting rule is emitted as a CSS class definition.
class ExportExcelToHtml
{
    static void Main()
    {
        // Load the source Excel workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Configure HTML save options to generate CSS classes for conditional formatting
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Export worksheet CSS separately so that conditional formatting rules become CSS classes
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Export all data (including styles) to ensure conditional formatting is included
        saveOptions.ExportDataOptions = HtmlExportDataOptions.All;

        // Keep all generated CSS (do not exclude unused styles) so conditional formatting CSS is retained
        saveOptions.ExcludeUnusedStyles = false;

        // Save the workbook as an HTML file
        string outputPath = "output.html";
        workbook.Save(outputPath, saveOptions);
    }
}
