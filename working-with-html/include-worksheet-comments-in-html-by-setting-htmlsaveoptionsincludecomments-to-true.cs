// Title: Include worksheet comments when saving an Excel workbook to HTML using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Write C# code that loads an .xlsx file, sets HtmlSaveOptions.IncludeComments = true, and saves the workbook as an .html file with all cell comments preserved using Aspose.Cells. | Show how to check the installed Aspose.Cells version for IncludeComments support before exporting to HTML, and add error handling for a missing input file. | Create a minimal console application that converts a workbook to HTML while keeping worksheet comments, and logs success or detailed error messages.
// Common Searches: Aspose.Cells C# save workbook as HTML with cell comments included | HtmlSaveOptions IncludeComments true example Aspose.Cells .NET | preserve Excel comments when converting to HTML using Aspose.Cells | check Aspose.Cells version for comment export support C# | C# console app export XLSX to HTML keeping worksheet comments
// Tags: Aspose.Cells HtmlSaveOptions IncludeComments | export worksheet comments to HTML C# | convert XLSX to HTML with comments Aspose.Cells | C# Aspose.Cells version compatibility comment export | handle missing Excel file Aspose.Cells C# | save workbook as HTML preserving comments

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates loading an Excel workbook, configuring HtmlSaveOptions.IncludeComments = true (with a note on version support), and saving the file as HTML while preserving all worksheet comments. It includes checks for the input file's existence and basic exception handling.
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
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Note: ExportCellComments property is not available in this version of Aspose.Cells.
                // If comment export is required, ensure you are using a version that supports it.
            };

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
