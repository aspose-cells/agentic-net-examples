// Title: Include hidden worksheet data when converting an Excel workbook to HTML with Aspose.Cells in C#
// AI Prompts: Write C# code that builds an Excel workbook, hides a worksheet, populates cells, enables HtmlSaveOptions.ExportHiddenWorksheet, and saves the result as an HTML file. | Show how to set up HtmlSaveOptions in Aspose.Cells so that hidden worksheets are rendered in the generated HTML output. | Create a C# example that converts several hidden worksheets into one HTML document using Aspose.Cells.
// Common Searches: Aspose.Cells C# export hidden sheets to HTML example | How to include hidden worksheet content in HTML output using Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet true C# sample code | Convert Excel to HTML with hidden worksheets visible Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# hidden worksheet HTML export | include hidden sheets in HTML conversion Aspose.Cells | Excel to HTML conversion with hidden worksheets C# | HtmlSaveOptions hide worksheet handling

using Aspose.Cells;
using System;

// The sample creates a workbook, adds a hidden worksheet named "HiddenSheet", writes values to cells A1 and B2, sets HtmlSaveOptions.ExportHiddenWorksheet = true, and saves the workbook as "ExportedWithHidden.html", causing the hidden sheet's content to appear in the generated HTML.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet and hide it
        int hiddenSheetIndex = workbook.Worksheets.Add();
        Worksheet hiddenSheet = workbook.Worksheets[hiddenSheetIndex];
        hiddenSheet.Name = "HiddenSheet";
        hiddenSheet.IsVisible = false; // Hide the worksheet

        // Populate hidden worksheet with some data
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");
        hiddenSheet.Cells["B2"].PutValue(123);

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = true;

        // Export the workbook to HTML
        workbook.Save("ExportedWithHidden.html", htmlOptions);
    }
}
