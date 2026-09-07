// Title: Export an Excel workbook to HTML with locale-specific date formatting using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, assigns Workbook.Settings.CultureInfo to a target locale (e.g., fr-FR), and saves the workbook as HTML so that all date cells appear in the locale's format. | Update existing Aspose.Cells HTML export logic to verify the source file, apply a custom CultureInfo, and implement robust exception handling while preserving regional date formatting in the generated HTML.
// Common Searches: how to export Excel to HTML with culture-specific dates using Aspose.Cells .NET | Aspose.Cells C# set workbook cultureinfo before HtmlSaveOptions | preserve regional date format when converting .xlsx to .html with Aspose.Cells | C# Aspose.Cells export to HTML respecting workbook locale settings | set workbook Settings.CultureInfo for date formatting in HTML output Aspose.Cells
// Tags: aspocells export excel to html with cultureinfo | c# workbook settings cultureinfo date formatting | htmlsaveoptions locale based date rendering | aspocells missing workbook file handling | regional date formatting aspocells html export

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// Loads 'input.xlsx', sets Workbook.Settings.CultureInfo to 'en-US', and saves it as 'output.html' via HtmlSaveOptions, ensuring dates are formatted according to the workbook's locale; includes file‑existence verification and exception handling.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputFile);

            // Set the workbook's locale (culture) to control date formatting.
            // Replace "en-US" with the desired locale identifier.
            workbook.Settings.CultureInfo = new CultureInfo("en-US");

            // Configure HTML save options. No explicit ExportDateTimeFormat property is required;
            // the workbook's CultureInfo setting controls date rendering.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file using the configured options.
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during loading or saving.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
