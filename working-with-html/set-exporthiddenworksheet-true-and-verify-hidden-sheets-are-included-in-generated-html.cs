// Title: Export hidden worksheets to HTML using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export hidden worksheets to HTML | How to include hidden sheets when saving Excel as HTML using Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet property example in .NET | Save workbook as HTML and retain hidden worksheets with Aspose.Cells | C# generate HTML from Excel that contains hidden worksheets
// Tags: Aspose.Cells HTML hidden worksheet export | HtmlSaveOptions ExportHiddenWorksheet C# | Include hidden sheets in HTML output Aspose.Cells | Export hidden worksheets to HTML .NET | Save workbook as HTML with hidden sheets C#

using Aspose.Cells;
using System;

// The sample loads an Excel workbook, hides a worksheet, sets HtmlSaveOptions.ExportHiddenWorksheet to true, and saves the workbook as HTML, ensuring that hidden worksheets are included in the generated HTML file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Example: ensure at least one worksheet is hidden
        // (In real scenarios the workbook may already contain hidden sheets)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.IsVisible = false; // Hide the first worksheet

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = true; // Export hidden sheets

        // Save the workbook as HTML; hidden sheets will be part of the output
        workbook.Save("output.html", htmlOptions);
    }
}
