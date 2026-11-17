// Title: Export hidden Excel worksheets to HTML with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells include hidden sheets when saving workbook as HTML | How to export hidden worksheets to HTML using HtmlSaveOptions in .NET | Aspose.Cells HTML export hidden worksheet property example | Save Excel file as HTML with hidden worksheets using Aspose.Cells C# code
// Tags: HtmlSaveOptions ExportHiddenWorksheet property | Aspose.Cells hidden worksheet HTML export | C# save Excel as HTML with hidden sheets | include hidden worksheets in HTML output .NET | Aspose.Cells HTMLSaveOptions hidden sheets

using Aspose.Cells;
using System;

// The example loads an Excel workbook, sets HtmlSaveOptions.ExportHiddenWorksheet to true, and saves the workbook as an HTML file, ensuring that any hidden worksheets are rendered in the generated HTML output.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = true; // Include hidden sheets in the exported HTML

        // Export the workbook to HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
