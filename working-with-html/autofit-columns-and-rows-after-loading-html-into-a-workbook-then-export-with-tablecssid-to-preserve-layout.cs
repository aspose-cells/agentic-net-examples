// Title: Auto‑Fit Columns & Rows When Importing HTML and Preserve Layout with TableCssId – Aspose.Cells for .NET
// Description: Shows how to load an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions.AutoFitColsAndRows, then export it back to HTML with HtmlSaveOptions.TableCssId so the original table styling is kept via CSS.
// Keywords: Aspose.Cells | C# | .NET | HtmlLoadOptions AutoFitColsAndRows | HtmlSaveOptions TableCssId | HTML to Excel conversion | Excel to HTML export | preserve table layout | auto‑fit columns rows | custom CSS id for HTML table
// Common Searches: Aspose.Cells auto fit columns rows when loading HTML | How to use TableCssId in HtmlSaveOptions | Preserve HTML table styling after Excel conversion | Load HTML into workbook and export with custom CSS selector | C# Aspose.Cells import HTML auto‑fit
// Developer Intent: Import an HTML document, automatically adjust its column widths and row heights, and save it as HTML with a designated CSS id to retain the original layout.
// Use Cases: Convert web‑based HTML reports to Excel, auto‑fit the cells, and generate HTML output that matches the site’s CSS framework. | Process user‑submitted HTML spreadsheets, ensure proper sizing, and re‑export with a custom TableCssId for seamless integration into existing web pages. | Migrate legacy HTML tables to Excel for data manipulation, then publish them back to HTML while preserving the original visual styling.
// AI Prompts: Provide C# code that loads an HTML file into an Aspose.Cells Workbook with AutoFitColsAndRows enabled and saves it as HTML using a custom TableCssId. | Explain how HtmlSaveOptions.TableCssId influences the generated HTML and how it works together with AutoFitColsAndRows to keep the table layout intact. | Show a step‑by‑step example of converting HTML to Excel, auto‑fitting rows and columns, then exporting back to HTML with a specific CSS selector for the table.

using System;
using Aspose.Cells;

// Shows how to load an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions.AutoFitColsAndRows, then export it back to HTML with HtmlSaveOptions.TableCssId so the original table styling is kept via CSS.
class Program
{
    static void Main()
    {
        // Load HTML with auto‑fit enabled for both columns and rows
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.AutoFitColsAndRows = true;               // Auto‑fit during import

        // Load the HTML file into a workbook using the specified options
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Prepare HTML save options and set a TableCssId to keep layout styling
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "custom-table-style";       // Prefix for CSS selectors in the generated table

        // Save the workbook as HTML with the configured options
        workbook.Save("output.html", saveOptions);
    }
}
