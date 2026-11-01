// Title: Export hidden Excel worksheets to HTML using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML while preserving hidden worksheets with Aspose.Cells. | Show how to enable HtmlSaveOptions.ExportHiddenWorksheet for including hidden sheets in an HTML export using Aspose.Cells .NET. | Provide a step‑by‑step example of exporting a workbook to HTML with hidden sheets included via the Aspose.Cells API.
// Common Searches: Aspose.Cells .NET example for exporting hidden worksheets to HTML | How to keep hidden Excel sheets in HTML output with C# Aspose.Cells | Using HtmlSaveOptions ExportHiddenWorksheet property in C# | Save workbook as HTML and include hidden sheets using Aspose.Cells | C# Aspose.Cells export all worksheets, hidden and visible, to HTML
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# hidden worksheet HTML export | include hidden sheets in HTML output Aspose.Cells | export Excel to HTML with hidden worksheets .NET | HTML export hidden worksheets Aspose.Cells

using System;
using Aspose.Cells;

// Loads an Excel workbook, sets HtmlSaveOptions.ExportHiddenWorksheet to true, and saves the workbook as an HTML file, ensuring hidden worksheets are included in the generated HTML.
class ExportHiddenWorksheetsToHtml
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = true; // Include hidden sheets in the export

        // Export the workbook to HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
