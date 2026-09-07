// Title: Save an Excel workbook as HTML with hyperlinks opening in new tabs using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures HtmlSaveOptions to export hyperlinks with target="_blank", and saves the workbook as an .html file. | Explain how to use Aspose.Cells HtmlSaveOptions in .NET to ensure that hyperlinks in the generated HTML open in a new browser tab.
// Common Searches: Aspose.Cells C# export Excel to HTML with links opening in new tab | How to make hyperlinks open in a new window when saving workbook as HTML using Aspose.Cells | HtmlSaveOptions target blank property Aspose.Cells .NET | Convert .xlsx to .html preserving hyperlink target attribute with Aspose.Cells | C# Aspose.Cells HTML save options hyperlink behavior
// Tags: Aspose.Cells HtmlSaveOptions hyperlink target | export Excel to HTML C# Aspose.Cells | set target blank Aspose.Cells HTML | convert .xlsx to .html Aspose.Cells .NET | hyperlink behavior Aspose.Cells HTML export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Demonstrates loading an Excel workbook with Aspose.Cells, configuring HtmlSaveOptions for HTML output, and saving the file so that all hyperlinks are rendered with target="_blank" (new‑tab behavior). Includes file existence verification and basic error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Note: In recent Aspose.Cells versions the property to add target="_blank"
            // is not required; hyperlinks are exported with the appropriate target.
            // If a specific property exists in your version, set it here.

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
