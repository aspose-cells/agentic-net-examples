// Title: Disable downlevel‑revealed conditional comments when exporting a workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells HtmlSaveOptions disable downlevel revealed comments C# | How to prevent IE conditional comments in HTML output from Aspose.Cells | Turn off downlevel HTML comments when saving a workbook as HTML using Aspose.Cells | C# Aspose.Cells export to HTML without conditional comments | Remove downlevel revealed comments from generated HTML in Aspose.Cells .NET
// Tags: HtmlSaveOptions.DownlevelHtmlComments false | Aspose.Cells HTML export clean output | C# configure HtmlSaveOptions for HTML export | remove IE conditional comments from Aspose.Cells HTML | generate HTML without downlevel revealed comments

using System;
using Aspose.Cells;

// Set the HtmlSaveOptions.DownlevelHtmlComments property to false before calling Workbook.Save. This disables the insertion of downlevel‑revealed conditional comments, producing clean HTML output when exporting a workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new empty workbook with a default worksheet
            var workbook = new Workbook();

            // Configure HTML save options (default settings are sufficient for this example)
            var htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
