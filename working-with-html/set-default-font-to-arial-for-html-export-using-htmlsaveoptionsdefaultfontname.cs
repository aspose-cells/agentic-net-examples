// Title: Set the default font to Arial when exporting an Excel workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file, configures HtmlSaveOptions to use Arial as the default font, and saves the workbook as an HTML document. | Demonstrate how to apply a custom default font during HTML conversion by setting HtmlSaveOptions.DefaultFontName in Aspose.Cells. | Update existing Aspose.Cells export logic to change the default font for HTML output to a specified type such as Arial.
// Common Searches: Aspose.Cells C# set default font Arial for HTML export | How to specify a default font in HtmlSaveOptions when saving a workbook as HTML | HtmlSaveOptions.DefaultFontName property example Aspose.Cells | Export Excel to HTML with a specific font using Aspose.Cells .NET | Change default font for HTML output in Aspose.Cells library
// Tags: Aspose.Cells HtmlSaveOptions default font | C# export Excel to HTML with custom font | HtmlSaveOptions.DefaultFontName usage | HTML conversion default font Aspose.Cells | Set Arial as default font for HTML export

using Aspose.Cells;

// // Loads input.xlsx, sets HtmlSaveOptions.DefaultFontName to "Arial", and saves the workbook as output.html.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML export options to use Arial as the default font
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DefaultFontName = "Arial";

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
