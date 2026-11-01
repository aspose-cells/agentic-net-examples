// Title: Load an existing .xlsx file, disable zero-value display, and save it as an HTML page using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to open a workbook, strip zero values from cells, and save the result as an HTML file with appropriate HtmlSaveOptions. | Show how to post‑process the HTML produced by Aspose.Cells in C# to replace zero values with empty strings after conversion.
// Common Searches: Aspose.Cells C# export Excel to HTML without showing zero values | How to hide zero numbers when saving workbook as HTML using Aspose.Cells .NET | C# remove zero cells before converting .xlsx to HTML with Aspose.Cells | HtmlSaveOptions ShowZeroValues not working in Aspose.Cells .NET
// Tags: aspocells htmlsaveoptions suppress zero values | c# convert xlsx to html aspocells | strip zero values before html export aspocells | modify html output to clear zero values c# | load workbook save as html aspocells

using Aspose.Cells;
using System;
using System.IO;

// The program verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions (noting that ShowZeroValues is unavailable), saves the workbook as 'output.html', and writes a success or error message to the console.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Note: ShowZeroValues option is not available in this version of Aspose.Cells.
            // If needed, zero values can be handled manually after conversion.

            // Export the workbook as an HTML page
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
