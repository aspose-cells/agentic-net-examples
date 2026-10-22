// Title: Export Excel to responsive HTML without hidden worksheets using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML while excluding hidden worksheets and enabling column widths to scale responsively with Aspose.Cells. | Create a reusable C# method that takes source and destination paths and applies HtmlSaveOptions.ExportHiddenWorksheet = false and HtmlSaveOptions.WidthScalable = true for HTML conversion. | Show how to adjust existing Aspose.Cells HTML export settings to disable hidden sheet output and activate responsive column scaling for web‑friendly pages.
// Common Searches: Aspose.Cells C# prevent hidden worksheets from appearing in HTML export | Enable responsive column widths when saving Excel as HTML with Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet false example in .NET | How to use WidthScalable property for mobile‑friendly HTML output from Aspose.Cells | Convert XLSX to responsive HTML using Aspose.Cells settings
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | Aspose.Cells WidthScalable responsive HTML | C# export Excel to HTML without hidden sheets | responsive column scaling Aspose.Cells | HTML conversion of XLSX with Aspose.Cells

using Aspose.Cells;
using System;

// // Load an Excel workbook, configure HtmlSaveOptions to skip hidden worksheets and enable scalable column widths, then save the file as responsive HTML.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set up HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportHiddenWorksheet = false; // Do not export hidden worksheets
        htmlOptions.WidthScalable = true; // Enable responsive column widths for HTML

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
