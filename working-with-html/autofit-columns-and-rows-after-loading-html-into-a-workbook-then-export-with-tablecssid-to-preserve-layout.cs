// Title: C# – Auto‑fit columns and rows when loading HTML with Aspose.Cells and save using a custom TableCssId
// Description: Loads an HTML file into an Aspose.Cells Workbook with AutoFitColsAndRows enabled, then exports the workbook back to HTML while assigning a TableCssId for targeted CSS styling. The approach keeps the original table layout and simplifies post‑processing in web applications.
// Keywords: Aspose.Cells HtmlLoadOptions AutoFitColsAndRows | Aspose.Cells HtmlSaveOptions TableCssId | C# auto fit columns rows HTML | preserve HTML table layout Aspose | .NET export workbook to HTML custom CSS id
// Common Searches: Aspose.Cells auto fit columns rows when loading HTML | Set TableCssId in Aspose.Cells HTML export | Preserve HTML table layout after conversion with Aspose.Cells | C# load HTML workbook auto‑fit and save with custom CSS id
// Developer Intent: Load an HTML document into a workbook, automatically adjust column widths and row heights, and save the workbook as HTML with a user‑defined TableCssId for precise CSS targeting.
// Use Cases: Render uploaded HTML reports in a web portal, auto‑fit the grid, and apply a consistent stylesheet via a custom TableCssId. | Convert HTML tables to Excel, retain the original layout through auto‑fit, then re‑export to HTML for further web‑based manipulation. | Batch‑process a collection of HTML files, applying auto‑fit on load and assigning unique TableCssIds to each output for uniform styling across all pages.
// AI Prompts: Write C# code that uses Aspose.Cells to load an HTML file, enable AutoFitColsAndRows, and save it as HTML with a specified TableCssId. | Explain the interaction between HtmlLoadOptions.AutoFitColsAndRows and HtmlSaveOptions.TableCssId in preserving table layout. | Provide a step‑by‑step guide for batch converting a folder of HTML files to HTML with auto‑fit and custom TableCssId using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells Workbook with AutoFitColsAndRows enabled, then exports the workbook back to HTML while assigning a TableCssId for targeted CSS styling. The approach keeps the original table layout and simplifies post‑processing in web applications.
class Program
{
    static void Main()
    {
        // Load HTML file with auto‑fit enabled for columns and rows
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.AutoFitColsAndRows = true;               // Auto‑fit during load

        // Replace "input.html" with the path to your source HTML file
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Configure HTML save options to use a custom TableCssId
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "custom-table-style";       // Prefix for CSS selectors inside the table

        // Save the workbook as HTML preserving the layout
        // Replace "output.html" with the desired output path
        workbook.Save("output.html", saveOptions);
    }
}
