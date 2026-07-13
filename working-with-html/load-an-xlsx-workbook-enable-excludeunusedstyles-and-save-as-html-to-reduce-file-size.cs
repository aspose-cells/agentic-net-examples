using System;
using Aspose.Cells; // Aspose.Cells library

// Author: Aspose.Cells .NET example – load XLSX, enable ExcludeUnusedStyles, save as HTML
class Program
{
    static void Main()
    {
        // Load the source workbook (XLSX format)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Exclude styles that are not used in the workbook to reduce HTML size.
            // The default is true, but we set it explicitly for clarity.
            ExcludeUnusedStyles = true
        };

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}