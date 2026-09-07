// Title: Include hidden worksheets when exporting Excel to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, sets HtmlSaveOptions.ExportHiddenWorksheet to true, and saves the workbook as an HTML file using Aspose.Cells. | Explain how to configure Aspose.Cells HtmlSaveOptions to retain hidden sheets during HTML conversion in a .NET application.
// Common Searches: Aspose.Cells C# export hidden worksheets to HTML example | How to keep hidden Excel sheets in HTML output using Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet true C# sample code | Saving workbook as HTML with hidden sheets included Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# include hidden sheets in HTML export | Aspose.Cells preserve hidden worksheets during HTML conversion | Workbook.Save with HtmlSaveOptions for hidden worksheets | Excel to HTML conversion retaining hidden sheets

using System;
using Aspose.Cells;

// The snippet loads an existing Excel workbook, enables the ExportHiddenWorksheet flag in HtmlSaveOptions, and saves the workbook as HTML, ensuring that any hidden worksheets are rendered in the output.
class ExportHiddenWorksheetsToHtml
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportHiddenWorksheet = true; // Include hidden sheets in the HTML output

        // Export the workbook to HTML (replace with your desired output path)
        workbook.Save("output.html", saveOptions);
    }
}
