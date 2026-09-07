// Title: Convert an Excel .xlsx file to HTML in C# with Aspose.Cells while rendering formulas as calculated values
// AI Prompts: Generate C# code that loads a workbook using Aspose.Cells, configures HtmlSaveOptions to output evaluated cell values, and saves the result as an HTML file. | Show how to verify the existence of an input .xlsx file, handle possible errors, and export it to HTML with Aspose.Cells so that formula cells appear as their computed results. | Demonstrate setting Aspose.Cells HtmlSaveOptions to suppress formula strings and produce static HTML that displays only the calculated values.
// Common Searches: aspnet convert xlsx to html using aspose.cells without showing formulas | c# export excel workbook to html with evaluated formula results | how to hide formula strings when saving Excel as HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions preserve calculated values in HTML output | save workbook as html showing only values not formulas c#
// Tags: Aspose.Cells HtmlSaveOptions export evaluated values | C# convert xlsx to html Aspose.Cells | Excel to HTML conversion without formula strings | save workbook as html calculated cell values | Aspose.Cells HTML rendering of formulas as values

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A C# example that checks for an input.xlsx file, loads it with Aspose.Cells, uses HtmlSaveOptions (which defaults to writing evaluated values) to convert the workbook to output.html, and includes basic error handling for missing files and runtime exceptions.
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

            // Configure HTML save options (default exports evaluated values)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
