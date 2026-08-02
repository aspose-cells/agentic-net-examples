// Title: Export a Worksheet to HTML with Styles and Merged Cells using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, merges A1:B1, applies custom styling, and uses HtmlSaveOptions to export only the active sheet to an HTML file. The export keeps all cell formatting, merged‑area validation, and embeds CSS inline, saving the result to the desktop.
// Keywords: Aspose.Cells | C# | HTML export | preserve cell styles | merged cells | HtmlSaveOptions | ExportActiveWorksheetOnly | ValidateMergedAreas | Excel to HTML | Aspose.Cells for .NET sample | GitHub example
// Common Searches: Aspose.Cells export worksheet to HTML C# | keep merged cells when saving Excel as HTML | HtmlSaveOptions preserve formatting Aspose | export single worksheet to HTML with styles | C# code to convert Excel to HTML with CSS inline
// Developer Intent: Generate an HTML representation of a worksheet that retains all visual formatting and merged‑cell structures.
// Use Cases: Produce a web‑ready report that mirrors the Excel layout, including merged headers and custom colors. | Display a single worksheet on a portal without losing styling, using a self‑contained HTML file. | Create an email attachment in HTML format that preserves the original spreadsheet appearance.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet to HTML, preserving merged cells and all styles. | Show how to configure HtmlSaveOptions for inline CSS, active‑sheet export, and merged‑area validation. | Explain steps to style cells before exporting to HTML with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// C# example that creates a workbook, merges A1:B1, applies custom styling, and uses HtmlSaveOptions to export only the active sheet to an HTML file. The export keeps all cell formatting, merged‑area validation, and embeds CSS inline, saving the result to the desktop.
class ExportWorksheetToHtml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sample";

        // Populate some data
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Data 1");
        sheet.Cells["B2"].PutValue(123);
        sheet.Cells["A3"].PutValue("Data 2");
        sheet.Cells["B3"].PutValue(456);

        // Merge cells A1:B1 to create a merged header
        sheet.Cells.Merge(0, 0, 1, 2); // row 0, column 0, rows 1, columns 2

        // Apply style to the merged header cell
        Style headerStyle = sheet.Cells["A1"].GetStyle();
        headerStyle.Font.Name = "Arial";
        headerStyle.Font.Size = 14;
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = Color.LightBlue;
        headerStyle.Pattern = BackgroundType.Solid;
        sheet.Cells["A1"].SetStyle(headerStyle);

        // Configure HTML save options to preserve styles and merged cells
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportActiveWorksheetOnly = true;          // Export only the active sheet
        saveOptions.ExportWorksheetCSSSeparately = false;     // Keep CSS in the same file
        saveOptions.ValidateMergedAreas = true;               // Ensure merged areas are validated
        saveOptions.ExportDataOptions = HtmlExportDataOptions.All; // Export all data including styles

        // Define output HTML file path (e.g., Desktop)
        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "WorksheetExport.html");

        // Save the workbook as HTML
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Worksheet exported to HTML at: " + outputPath);
    }
}
