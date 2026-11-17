// Title: Export Excel worksheet headings as separate HTML tables with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export worksheet headings as individual HTML tables | How to enable ExportHeadings in HtmlSaveOptions for Excel to HTML conversion | Save Excel workbook to HTML with headings in separate tables using Aspose.Cells | C# code example for HtmlSaveOptions.ExportHeadings property in Aspose.Cells | Generate HTML output with distinct heading tables from an .xlsx file in .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportHeadings | Excel to HTML conversion with separate heading tables | C# Aspose.Cells export headings HTML | HTML output with worksheet headings Aspose.Cells | Separate HTML tables for Excel column headers

using System;
using Aspose.Cells;

// Loads an Excel workbook, sets HtmlSaveOptions.ExportHeadings to true, and saves the file as HTML where each worksheet's headings are emitted as independent HTML tables.
class ExportHeadingsToHtml
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to export headings as separate tables
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHeadings = true; // Enable exporting of worksheet headings

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
