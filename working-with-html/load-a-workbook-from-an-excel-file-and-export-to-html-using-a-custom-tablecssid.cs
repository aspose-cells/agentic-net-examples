// Title: Export an Excel workbook to HTML with a custom TableCssId using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as HTML while assigning a custom TableCssId to the generated table. | Show how to configure HtmlSaveOptions in Aspose.Cells to specify a CSS identifier for the HTML table when converting a workbook.
// Common Searches: how to set TableCssId in Aspose.Cells HtmlSaveOptions C# | Aspose.Cells export Excel to HTML with custom table id example | C# convert .xlsx to HTML with specific table CSS identifier using Aspose.Cells | customize HTML table id when saving workbook as HTML Aspose.Cells | Aspose.Cells HTML export custom CSS id for table
// Tags: Aspose.Cells HtmlSaveOptions custom TableCssId | C# HTML export with table CSS identifier | custom CSS id for HTML table Aspose.Cells | Excel to HTML conversion with table id | Aspose.Cells workbook HTML save with CSS id

using System;
using Aspose.Cells;

// Loads an Excel file, configures HtmlSaveOptions with a custom TableCssId, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options with a custom CSS ID for the generated table
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.TableCssId = "myCustomTableId";

        // Export the workbook to HTML using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
