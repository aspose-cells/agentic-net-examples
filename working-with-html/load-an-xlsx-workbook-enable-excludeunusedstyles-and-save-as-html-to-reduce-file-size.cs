// Title: Convert an XLSX workbook to HTML with Aspose.Cells for .NET while excluding unused styles to shrink output size
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, sets HtmlSaveOptions.ExcludeUnusedStyles = true, and saves the workbook as a compact HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions to prune unused CSS when exporting Excel to HTML in a .NET application. | Demonstrate reducing the HTML file size generated from an Excel workbook by disabling unused style generation with Aspose.Cells.
// Common Searches: convert excel to html minimal css Aspose.Cells | C# Aspose.Cells reduce html file size | how to export xlsx as html without extra styles | Aspose.Cells HtmlSaveOptions shrink html output | save workbook as html with compact css using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExcludeUnusedStyles | C# Excel to HTML conversion optimization | remove unused CSS Aspose.Cells | minimize HTML payload from workbook export | export XLSX as HTML with style pruning

using Aspose.Cells;
using System;

// This C# example loads an XLSX workbook, enables HtmlSaveOptions.ExcludeUnusedStyles to omit unused CSS, and saves the workbook as a smaller HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to exclude unused styles (reduces file size)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExcludeUnusedStyles = true;

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
