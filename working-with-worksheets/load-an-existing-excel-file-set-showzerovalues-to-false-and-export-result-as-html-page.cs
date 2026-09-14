// Title: Load an Excel workbook, disable zero-value display, and save it as an HTML page using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file with Aspose.Cells, sets HtmlSaveOptions.ShowZeroValues to false, and saves the workbook as an HTML file. | Provide a C# snippet that verifies the Excel file exists, configures HtmlSaveOptions to hide zero values, and exports the workbook to HTML with robust error handling. | Show how to customize Aspose.Cells HTML export to suppress zero values while keeping other default options.
// Common Searches: Aspose.Cells hide zero values when converting Excel to HTML in C# | C# export workbook to HTML without displaying zero cells using Aspose.Cells | How to set ShowZeroValues false in HtmlSaveOptions Aspose.Cells .NET | Check file existence before converting xlsx to html with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ShowZeroValues | C# convert Excel to HTML Aspose.Cells | disable zero values in HTML export Aspose.Cells | file existence check before Aspose.Cells workbook load | exception handling Aspose.Cells save to HTML

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an existing Excel file, optionally verifying its presence, configuring HtmlSaveOptions to suppress zero-value cells (ShowZeroValues = false), and saving the workbook as an HTML page while handling potential errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options (default settings)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export the workbook to an HTML page
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
