// Title: How to export hidden worksheets to HTML using Aspose.Cells HtmlSaveOptions in C#
// Common Searches: Aspose.Cells C# export hidden Excel sheets to HTML | Include hidden worksheets in HTML output with Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet example C# | Save workbook as HTML while preserving hidden sheet data Aspose | C# generate HTML from Excel that shows hidden sheet content
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | Aspose.Cells hidden worksheet HTML export | HTML conversion of hidden Excel sheets C# | preserve hidden sheet data during HTML export | generate HTML with visible hidden worksheets

using System;
using Aspose.Cells;

// Demonstrates loading an Excel workbook, setting HtmlSaveOptions.ExportHiddenWorksheet to true, and saving the file as HTML so that content from hidden worksheets is included in the generated HTML output.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportHiddenWorksheet = true; // Export content of hidden sheets

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
