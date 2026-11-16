// Title: Convert an XLSX workbook to HTML in C# with Aspose.Cells while excluding unused styles to shrink the output
// Common Searches: C# Aspose.Cells convert XLSX to HTML without unused CSS | How to reduce HTML file size when exporting Excel using Aspose.Cells | Aspose.Cells HtmlSaveOptions.ExcludeUnusedStyles example for .NET | Save Excel workbook as HTML with minimal styles in C# | Aspose.Cells HTML export size optimization technique
// Tags: Aspose.Cells HtmlSaveOptions ExcludeUnusedStyles | C# convert XLSX to HTML Aspose.Cells | reduce generated HTML size Aspose.Cells | Excel to HTML minimal CSS export | Aspose.Cells workbook.Save with HTML options

using Aspose.Cells;
using System;

// Loads an XLSX workbook, configures HtmlSaveOptions.ExcludeUnusedStyles to true to eliminate unused CSS, and saves the workbook as a compact HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to exclude unused styles
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExcludeUnusedStyles = true; // Reduces the size of the generated HTML

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
