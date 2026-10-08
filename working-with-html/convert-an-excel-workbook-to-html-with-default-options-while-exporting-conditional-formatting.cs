// Title: Convert an Excel workbook to HTML with default options while preserving conditional formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a .xlsx file, verifies the file exists, and saves it as an HTML document with Aspose.Cells default HtmlSaveOptions, ensuring conditional formatting is kept. | Create a try‑catch example that converts an Excel workbook to HTML with Aspose.Cells and logs any errors, confirming successful export.
// Common Searches: asp.net convert xlsx to html preserving conditional formatting with Aspose.Cells | c# Aspose.Cells export workbook to html using default save options | how to check if Excel file exists before using Aspose.Cells conversion | sample code for saving Excel as HTML and handling exceptions in C# | Aspose.Cells HtmlSaveOptions default behavior for conditional formatting
// Tags: Aspose.Cells HtmlSaveOptions default export | preserve conditional formatting when saving Excel as HTML C# | verify workbook file presence prior to conversion | exception handling for Aspose.Cells Save method C# | convert .xlsx to .html with Aspose.Cells API

using Aspose.Cells;
using System;
using System.IO;

// The example checks that input.xlsx exists, loads it into an Aspose.Cells Workbook, creates default HtmlSaveOptions (which automatically export conditional formatting), saves the workbook as output.html, and wraps the operation in a try‑catch block to report any errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputFile);

            // Create HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
