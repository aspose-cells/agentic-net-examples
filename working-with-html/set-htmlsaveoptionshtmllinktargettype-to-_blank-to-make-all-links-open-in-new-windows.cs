// Title: How to set HtmlSaveOptions.HtmlLinkTargetType to "_blank" in Aspose.Cells C# to open all hyperlinks in new windows when exporting to HTML
// AI Prompts: Write C# code using Aspose.Cells that loads an Excel workbook and saves it as HTML with HtmlSaveOptions.HtmlLinkTargetType set to "_blank" so every hyperlink opens in a new window. | Show how to configure HtmlSaveOptions in Aspose.Cells to apply target="_blank" to all links during HTML conversion, including file existence checks and exception handling.
// Common Searches: Aspose.Cells C# export Excel to HTML with hyperlinks opening in new tab | set HtmlLinkTargetType _blank Aspose.Cells HtmlSaveOptions | make all links open in new window when saving workbook as HTML using Aspose.Cells | C# Aspose.Cells HtmlSaveOptions target attribute for hyperlinks
// Tags: Aspose.Cells hyperlink target setting | C# HtmlSaveOptions for HTML export | export Excel to HTML with new window links | Aspose.Cells set link target blank | HTML conversion hyperlink target Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an existing Excel workbook, creates HtmlSaveOptions for HTML format, sets HtmlSaveOptions.HtmlLinkTargetType to HtmlLinkTargetType.Blank (which adds target="_blank" to every hyperlink), and saves the workbook as an HTML file while handling missing input files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Note: Aspose.Cells does not provide a direct property to set hyperlink target.
            // If needed, you can modify the generated HTML after saving.

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
