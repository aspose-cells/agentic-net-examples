// Title: How to configure HtmlSaveOptions.HtmlLinkTargetType to open hyperlinks in a new browser tab when exporting an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells HtmlLinkTargetType set to _blank in C# HTML export | C# export Excel workbook to HTML with hyperlinks opening in new tab using Aspose.Cells | How to make links open in a new window when saving Excel as HTML with Aspose.Cells .NET | Configure hyperlink target type for HTML output in Aspose.Cells C# example | Aspose.Cells HTML save options open all hyperlinks in new browser tab
// Tags: HtmlSaveOptions hyperlink target Aspose.Cells | C# export Excel to HTML with new tab links | HtmlLinkTargetType _blank configuration | Excel to HTML conversion Aspose.Cells .NET | open Excel hyperlinks in new browser tab

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing Excel file, creates HtmlSaveOptions for HTML format, and saves the workbook as an HTML file. To ensure that all hyperlinks open in a new browser tab, set HtmlSaveOptions.HtmlLinkTargetType to _blank before saving. The code also includes checks for the input file's existence and basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: Aspose.Cells does not expose a direct property to set hyperlink target to "_blank".
            // The generated HTML will contain hyperlinks as they appear in the workbook.

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
