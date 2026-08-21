// Title: Hide Columns C & D and Export Workbook While Preserving Hidden State with Aspose.Cells for .NET
// Description: Load an Excel file, hide columns C and D using the HideColumns method, save the workbook so the columns stay hidden, and export to HTML with HtmlSaveOptions configured to retain hidden columns.
// Keywords: Aspose.Cells | HideColumns | hide column C | hide column D | preserve hidden columns | export to HTML | HtmlSaveOptions | .NET Excel manipulation | Excel to HTML conversion | hidden column display type
// Common Searches: Aspose.Cells hide column C programmatically | keep hidden columns when saving Excel with C# | export Excel to HTML while retaining hidden columns | C# HideColumns multiple columns Aspose | HtmlSaveOptions hidden column setting example
// Developer Intent: Programmatically hide selected columns and ensure their hidden status is maintained in both the saved Excel file and the generated HTML output.
// Use Cases: Create a financial report where sensitive columns are hidden for viewers but remain in the source file, then share the report as HTML. | Prepare a template workbook, hide template‑specific columns before distributing it to clients, and keep those columns hidden when converting to web format. | Maintain data integrity while presenting an Excel sheet on a website, ensuring hidden columns stay concealed in the HTML view.
// AI Prompts: Generate C# code with Aspose.Cells that hides columns 2‑5 and saves the workbook as PDF while keeping the hidden column settings. | Show how to configure HtmlSaveOptions to show hidden rows instead of columns in an Aspose.Cells export. | Explain the steps to unhide columns that were previously hidden using the HideColumns method in Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel file, hide columns C and D using the HideColumns method, save the workbook so the columns stay hidden, and export to HTML with HtmlSaveOptions configured to retain hidden columns.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Hide columns C and D (zero‑based column indexes 2 and 3)
        // HideColumns(startColumn, totalColumns) hides a range of columns
        workbook.Worksheets[0].Cells.HideColumns(2, 2);

        // Save the workbook back to Excel format.
        // Hidden columns are retained automatically when saving in Excel formats.
        workbook.Save("output.xlsx");

        // If you need to export to HTML while preserving the hidden columns,
        // configure HtmlSaveOptions to keep hidden columns (default is Hidden).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            HiddenColDisplayType = HtmlHiddenColDisplayType.Hidden
        };
        workbook.Save("output.html", htmlOptions);
    }
}
