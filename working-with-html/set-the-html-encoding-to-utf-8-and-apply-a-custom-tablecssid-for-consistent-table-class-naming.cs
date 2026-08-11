// Title: Export Workbook to HTML with UTF‑8 Encoding and Custom TableCssId using Aspose.Cells for .NET
// Description: Step‑by‑step example showing how to save an Aspose.Cells workbook as HTML with UTF‑8 encoding and a custom TableCssId, ensuring proper character rendering and easy CSS styling.
// Keywords: Aspose.Cells HTML export | HtmlSaveOptions UTF-8 | TableCssId custom CSS | C# Aspose.Cells save as HTML | UTF-8 encoding Aspose.Cells | custom table CSS id Aspose | Excel to HTML Aspose.Cells
// Common Searches: Aspose.Cells set HTML encoding to UTF-8 | How to assign TableCssId in Aspose.Cells | HtmlSaveOptions custom CSS id example | Export Excel to HTML with UTF-8 using Aspose | Aspose.Cells C# HTML export styling
// Developer Intent: Configure HtmlSaveOptions to produce UTF‑8 encoded HTML and assign a specific TableCssId for consistent table styling when exporting a workbook.
// Use Cases: Publish multilingual spreadsheets on the web without character corruption. | Apply a predefined CSS rule set to the exported HTML table for brand‑consistent styling. | Combine proper encoding and custom CSS identifiers to meet accessibility and localization standards.
// AI Prompts: Show how to set HtmlSaveOptions.Encoding to UTF-8 and TableCssId to a custom value in Aspose.Cells for .NET. | Provide a complete C# example that saves a workbook as HTML with UTF-8 encoding and a custom TableCssId, and explain how to link the CSS file. | Explain how to modify the TableCssId after export or add multiple CSS classes to the generated HTML table using Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

// Step‑by‑step example showing how to save an Aspose.Cells workbook as HTML with UTF‑8 encoding and a custom TableCssId, ensuring proper character rendering and easy CSS styling.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);

        // Configure HTML save options: UTF-8 encoding and custom TableCssId
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.Encoding = Encoding.UTF8;          // Set HTML encoding to UTF-8
        saveOptions.TableCssId = "custom-table-style"; // Apply custom TableCssId

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", saveOptions);
    }
}
