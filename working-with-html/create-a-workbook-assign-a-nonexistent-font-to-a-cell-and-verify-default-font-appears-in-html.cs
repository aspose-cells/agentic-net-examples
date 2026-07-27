// Title: Aspose.Cells C# – Export Workbook to HTML with a Default Font for Missing Fonts
// Description: Creates a workbook, assigns a non‑existent font to a cell, sets HtmlSaveOptions.DefaultFontName (e.g., Arial) as a fallback, and saves the file as HTML so the missing font is automatically replaced.
// Keywords: Aspose.Cells HTML export | DefaultFontName | fallback font | missing font handling | C# Aspose.Cells | HtmlSaveOptions | font substitution | Excel to HTML conversion | nonexistent font
// Common Searches: Aspose.Cells set default font for HTML export | HtmlSaveOptions DefaultFontName example | How to handle missing fonts in Aspose.Cells HTML | C# export Excel to HTML with fallback font | Aspose.Cells replace unavailable font
// Developer Intent: Ensure that cells using fonts not installed on the target machine are rendered with a specified fallback font when the workbook is saved as HTML.
// Use Cases: Generate web‑ready reports from spreadsheets that contain custom or legacy fonts. | Create HTML previews for users on systems without the original fonts installed. | Automate email attachments where consistent HTML rendering is required across devices. | Build dashboards that display Excel data in browsers without needing additional font installations.
// AI Prompts: Write C# code using Aspose.Cells to apply a style with a potentially missing font and configure HtmlSaveOptions.DefaultFontName to Arial. | Explain how HtmlSaveOptions.DefaultFontName works and its impact on font fallback during HTML conversion. | Show how to detect installed fonts in .NET before assigning them to a cell style in Aspose.Cells. | Provide a step‑by‑step guide to export an Excel workbook to HTML with a fallback font using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, assigns a non‑existent font to a cell, sets HtmlSaveOptions.DefaultFontName (e.g., Arial) as a fallback, and saves the file as HTML so the missing font is automatically replaced.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample text to a cell
        worksheet.Cells["A1"].PutValue("Text with a missing font");

        // Create a style that uses a font name that does not exist on the system
        Style missingFontStyle = workbook.CreateStyle();
        missingFontStyle.Font.Name = "NonExistentFontXYZ";
        worksheet.Cells["A1"].SetStyle(missingFontStyle);

        // Configure HTML save options to use a known default font when the original font is missing
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DefaultFontName = "Arial";

        // Save the workbook to HTML; the cell will be rendered with the default font
        workbook.Save("output.html", htmlOptions);

        Console.WriteLine($"HTML saved. Missing font replaced by default font: {htmlOptions.DefaultFontName}");
    }
}
