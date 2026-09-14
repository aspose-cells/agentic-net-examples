// Title: Export an Excel workbook to HTML without generating CSS files using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a Workbook as HTML with CSS disabled using Aspose.Cells HtmlSaveOptions. | Show how to configure Aspose.Cells options to suppress external CSS creation during HTML export in a .NET application. | Provide a step‑by‑step example converting an .xlsx file to .html while preventing stylesheet generation with Aspose.Cells.
// Common Searches: Aspose.Cells disable CSS generation when saving workbook as HTML | C# example of turning off stylesheet output in Aspose.Cells HTML export | How to export Excel to HTML without external stylesheet in .NET | Prevent creation of CSS files during Aspose.Cells HTML export | Save workbook as HTML with no CSS using Aspose.Cells for C#
// Tags: Aspose.Cells HtmlSaveOptions.DisableCss | C# Excel to HTML conversion without stylesheet | Aspose.Cells suppress CSS output | HTML export without external CSS Aspose.Cells | save workbook as HTML no CSS Aspose.Cells

using Aspose.Cells;

// The sample loads an Excel workbook, creates HtmlSaveOptions with the DisableCss property set to true to stop CSS file creation, and saves the workbook as an HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to disable CSS generation
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.DisableCss = true; // Prevent creation of external CSS files

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
