// Title: Export Hidden Worksheets to HTML with HtmlCrossType.Cross – Aspose.Cells C# Example
// Description: C# sample that saves an Aspose.Cells workbook as HTML, includes hidden sheets (ExportHiddenWorksheet = true) and accelerates conversion of large workbooks using HtmlCrossType.Cross.
// Keywords: Aspose.Cells | C# | ExportHiddenWorksheet | HtmlCrossType.Cross | HTML export | hidden worksheet | large workbook performance | HtmlSaveOptions | Aspose.Cells HTML conversion | optimize HTML export
// Common Searches: Aspose.Cells export hidden worksheet to HTML | HtmlCrossType.Cross for large workbooks | How to include hidden sheets in HTML output Aspose.Cells | Improve HTML export speed Aspose.Cells C# | Export workbook as HTML with hidden sheets
// Developer Intent: Generate an HTML file from a workbook that contains hidden worksheets while maximizing performance for large files.
// Use Cases: Create a web‑ready report that shows data from both visible and hidden sheets for compliance audits. | Publish a massive Excel workbook on a website, ensuring hidden tabs are visible to end users. | Reduce conversion time when rendering large Excel files to HTML by using the Cross string type.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to HTML with ExportHiddenWorksheet enabled and HtmlCrossType.Cross, and describe the performance benefits. | Explain the internal mechanism of HtmlCrossType.Cross and when it should be chosen for large Excel files. | Show how to modify the example to export only specific hidden worksheets while leaving others excluded.

using System;
using Aspose.Cells;

// C# sample that saves an Aspose.Cells workbook as HTML, includes hidden sheets (ExportHiddenWorksheet = true) and accelerates conversion of large workbooks using HtmlCrossType.Cross.
class ExportHtmlExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add data to the first (visible) worksheet
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Visible Data");

        // Add a hidden worksheet with some data
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");
        hiddenSheet.IsVisible = false; // mark the sheet as hidden

        // Configure HTML save options:
        // - ExportHiddenWorksheet = true ensures hidden sheets are included in the output.
        // - HtmlCrossStringType = HtmlCrossType.Cross improves performance for large workbooks.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = true,
            HtmlCrossStringType = HtmlCrossType.Cross
        };

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
