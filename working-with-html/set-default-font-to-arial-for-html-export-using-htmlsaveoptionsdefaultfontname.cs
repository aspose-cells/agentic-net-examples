// Title: How to set Arial as the default font when exporting a workbook to HTML with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export workbook to HTML with Arial as default font | set default font for HTML output using HtmlSaveOptions in Aspose.Cells | C# code to change default font for HTML save options in Aspose.Cells | HtmlSaveOptions.DefaultFontName example for Excel to HTML conversion | configure default font for HTML export in Aspose.Cells library
// Tags: Aspose.Cells HtmlSaveOptions default font | C# export Excel to HTML with specific font | HtmlSaveOptions DefaultFontName usage | set Arial font for HTML conversion Aspose.Cells | Excel workbook HTML export configuration C#

using Aspose.Cells;
using Aspose.Cells.Saving;

// Loads an Excel workbook, sets HtmlSaveOptions.DefaultFontName to "Arial", and saves the workbook as an HTML file using those options.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options and set the default font to Arial
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.DefaultFontName = "Arial";

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
