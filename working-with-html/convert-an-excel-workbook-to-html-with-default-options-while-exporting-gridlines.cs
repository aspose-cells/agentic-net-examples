// Title: Convert an Excel workbook to HTML with gridlines using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file and saves it as an HTML document with gridlines enabled using Aspose.Cells. | Show how to set the ExportGridLines property on HtmlSaveOptions while converting a workbook to HTML in .NET.
// Common Searches: Aspose.Cells C# export workbook to HTML with gridlines | Enable gridlines when saving Excel as HTML using Aspose.Cells .NET | Default HtmlSaveOptions for HTML conversion in Aspose.Cells | How to preserve Excel gridlines in HTML output with Aspose.Cells | Saving .xlsx as .html with visible gridlines in C#
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | C# convert Excel to HTML with gridlines | Aspose.Cells default HTML conversion settings | Export Excel gridlines to HTML using .NET | Workbook.Save HTML Aspose.Cells example

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel file (input.xlsx) with Aspose.Cells, configures HtmlSaveOptions to export gridlines, and saves the workbook as an HTML file (output.html) using the default conversion settings.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options with default settings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Enable exporting of gridlines
        htmlOptions.ExportGridLines = true;

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
