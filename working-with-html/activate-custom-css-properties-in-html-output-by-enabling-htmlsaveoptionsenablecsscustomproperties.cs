// Title: How to enable CSS custom properties when exporting an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML with CSS custom properties enabled | EnableCssCustomProperties option in HtmlSaveOptions for .NET | Save workbook as HTML using Aspose.Cells while preserving CSS variables | C# code to activate CSS custom properties in HTML output from Excel | Aspose.Cells HTML export custom CSS variables example
// Tags: HtmlSaveOptions EnableCssCustomProperties | Aspose.Cells HTML export with CSS variables | C# Excel to HTML conversion using Aspose | custom CSS properties in generated HTML | configure Aspose.Cells HTML save options

using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel workbook, sets HtmlSaveOptions.EnableCssCustomProperties to true, and saves the file as HTML, producing output that utilizes CSS custom properties for styling.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to enable CSS custom properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.EnableCssCustomProperties = true; // Activate custom CSS properties

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
