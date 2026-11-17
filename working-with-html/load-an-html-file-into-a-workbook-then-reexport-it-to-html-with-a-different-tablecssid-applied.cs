// Title: C# – Load an HTML file into an Aspose.Cells Workbook and save it with a custom TableCssId
// Common Searches: how to change the HTML table id when exporting a workbook with Aspose.Cells in C# | Aspose.Cells load html file and set TableCssId in HtmlSaveOptions | C# code to read HTML into a Workbook and save with a new CSS identifier | customizing table CSS id for HTML output using Aspose.Cells HtmlSaveOptions | example of using HtmlSaveOptions.TableCssId property in Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | load HTML into Aspose.Cells workbook C# | export workbook to HTML with custom table id | custom CSS identifier for Aspose.Cells HTML output | C# Aspose.Cells HTML conversion TableCssId

using System;
using Aspose.Cells;

// The example loads an HTML file into an Aspose.Cells Workbook, sets HtmlSaveOptions.TableCssId to a new identifier, and saves the workbook back to HTML using the specified options.
class Program
{
    static void Main()
    {
        // Load the source HTML file into a workbook.
        // Aspose.Cells can directly read HTML files.
        Workbook workbook = new Workbook("input.html");

        // Configure HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        // Apply a new CSS ID to the generated HTML table.
        saveOptions.TableCssId = "newTableId";

        // Export the workbook back to HTML using the specified options.
        workbook.Save("output.html", saveOptions);
    }
}
