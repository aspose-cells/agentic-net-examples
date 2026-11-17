// Title: How to include worksheet comments when converting an Excel workbook to HTML with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export Excel worksheet comments to HTML | Set HtmlSaveOptions.IncludeComments true example Aspose.Cells | Save Excel file as HTML preserving cell notes using Aspose.Cells .NET | Aspose.Cells HTML conversion include comments option C# code sample | Convert .xlsx to .html with comments using Aspose.Cells library
// Tags: HtmlSaveOptions IncludeComments property | Excel to HTML conversion preserving comments | Aspose.Cells HTML export settings C# | Workbook save as HTML with cell notes | C# Aspose.Cells preserve worksheet comments

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, configures HtmlSaveOptions (which includes worksheet comments when IncludeComments is true), and saves the workbook as an HTML file, demonstrating how to retain cell comments in the HTML output while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default behavior includes cell comments)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
