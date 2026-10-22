// Title: Convert an Excel workbook to HTML with default options and include hidden worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file and saves it as an HTML document using Aspose.Cells with default HtmlSaveOptions while preserving hidden worksheets. | Demonstrate how to enable hidden worksheet export when converting a workbook to HTML with Aspose.Cells for .NET. | Provide a minimal example that converts an Excel workbook to HTML, using the default save options and ensuring hidden sheets appear in the output.
// Common Searches: Aspose.Cells C# export hidden worksheets to HTML | how to save Excel as HTML with default HtmlSaveOptions using Aspose.Cells | include hidden sheets in HTML output Aspose.Cells .NET | HtmlSaveOptions ExportHiddenWorksheet usage example C# | convert .xlsx to .html preserving hidden worksheets Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | Excel workbook to HTML conversion C# | preserve hidden sheets Aspose.Cells HTML output | default HTML save options Aspose.Cells | C# Aspose.Cells HTML export sample

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads 'input.xlsx', creates HtmlSaveOptions with default settings, enables ExportHiddenWorksheet, and saves the workbook as 'output.html' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options with default settings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export hidden worksheets as part of the HTML output
        htmlOptions.ExportHiddenWorksheet = true;

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
