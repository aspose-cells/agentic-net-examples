// Title: Export Excel workbook to responsive HTML while excluding hidden worksheets using Aspose.Cells for .NET
// Common Searches: how to export an Excel workbook to HTML without hidden sheets using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet false example in .NET | C# generate responsive HTML from Excel with column width scaling using Aspose.Cells | exclude hidden worksheets from HTML export Aspose.Cells WidthScalable true | make HTML output responsive column widths Aspose.Cells C#
// Tags: HtmlSaveOptions ExportHiddenWorksheet false | HtmlSaveOptions WidthScalable true | responsive column widths Aspose.Cells HTML | exclude hidden worksheets Aspose.Cells HTML export | C# Aspose.Cells HTML conversion

using Aspose.Cells;
using System;

// The sample creates a workbook with a visible and a hidden worksheet, sets the hidden sheet's visibility to false, configures HtmlSaveOptions to skip hidden worksheets (ExportHiddenWorksheet = false) and enable responsive column widths (WidthScalable = true), then saves the workbook as a responsive HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add data to the first (visible) worksheet
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Hello");
        visibleSheet.Cells["B1"].PutValue("World");

        // Add a hidden worksheet (will not be exported)
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.IsVisible = false;
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");

        // Configure HTML export options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = false; // exclude hidden worksheets
        htmlOptions.WidthScalable = true;          // enable responsive column widths

        // Export the workbook to HTML
        workbook.Save("output.html", htmlOptions);
    }
}
