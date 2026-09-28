// Title: Export worksheet headings as individual HTML tables with Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Generate C# code that loads an .xlsx workbook, sets HtmlSaveOptions.ExportHeadings = true, and saves it as HTML where each heading appears in its own table. | Show how to export only the heading rows of a specific worksheet to separate HTML tables using Aspose.Cells HtmlSaveOptions in C#. | Adapt the example to specify a custom output folder and file name while keeping ExportHeadings enabled for HTML conversion.
// Common Searches: Aspose.Cells C# export worksheet headings as separate HTML tables | Enable ExportHeadings in HtmlSaveOptions when converting Excel to HTML | Save Excel file to HTML with headings in distinct tables using Aspose.Cells .NET | Export only header rows of an Excel sheet to HTML with Aspose.Cells | C# Aspose.Cells HtmlSaveOptions ExportHeadings example
// Tags: HtmlSaveOptions ExportHeadings | export worksheet headings as html tables | aspocells excel to html conversion | c# html save options for excel headings | separate heading tables in html output

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Loads an Excel workbook, enables HtmlSaveOptions.ExportHeadings, and saves it as HTML where each worksheet heading is rendered in its own table.
class ExportHeadingsToHtml
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to export headings as separate tables
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.ExportHeadings = true; // Enable exporting of worksheet headings

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", saveOptions);
    }
}
