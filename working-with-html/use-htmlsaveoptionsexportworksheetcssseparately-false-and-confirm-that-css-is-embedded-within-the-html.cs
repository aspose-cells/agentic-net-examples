// Title: How to embed worksheet CSS directly in HTML when exporting an Excel workbook with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export workbook to HTML with inline CSS | prevent separate CSS file when saving Excel as HTML using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportWorksheetCSSSeparately false example code | verify generated HTML contains <style> tag after Aspose.Cells export | embed worksheet styles in HTML output with Aspose.Cells C#
// Tags: Aspose.Cells embed CSS in HTML output | HtmlSaveOptions ExportWorksheetCSSSeparately false | C# save workbook as HTML with inline stylesheet | verify inline <style> tag in Aspose.Cells HTML export | Excel to HTML conversion with embedded CSS using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, applies formatting, sets HtmlSaveOptions.ExportWorksheetCSSSeparately to false so the CSS is embedded, saves the workbook as HTML, reads the file to check for a <style> tag, and outputs the file path along with a boolean indicating whether the CSS was embedded.
class HtmlExportDemo
{
    static void Main()
    {
        // Create a new workbook and add some data with formatting
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sample";

        // Populate cells
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Apply some formatting to generate CSS
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = System.Drawing.Color.Blue;
        sheet.Cells["A1"].SetStyle(headerStyle);

        // Configure HTML save options to embed CSS within the HTML
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = false; // Embed CSS

        // Define output HTML file path
        string htmlPath = "ExportedWorkbook.html";

        // Save the workbook as HTML with the specified options
        workbook.Save(htmlPath, saveOptions);

        // Verify that CSS is embedded by checking for a <style> tag in the HTML file
        string htmlContent = File.ReadAllText(htmlPath);
        bool cssEmbedded = htmlContent.Contains("<style");

        Console.WriteLine($"HTML file saved to: {Path.GetFullPath(htmlPath)}");
        Console.WriteLine($"CSS embedded within HTML: {cssEmbedded}");
    }
}
