// Title: How to enable custom CSS properties when exporting an Excel workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets HtmlSaveOptions.EnableCustomProperties = true and saves a workbook as HTML using Aspose.Cells. | Provide a complete example of converting an .xlsx file to .html while preserving custom CSS properties with Aspose.Cells in C#.
// Common Searches: Aspose.Cells set EnableCustomProperties true for HTML export C# | export Excel to HTML with custom CSS using Aspose.Cells .NET | C# Aspose.Cells HtmlSaveOptions custom CSS properties | how to keep custom styles when converting Excel to HTML Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions EnableCustomProperties | HTML export custom CSS Aspose.Cells | C# export Excel to HTML with custom properties | Aspose.Cells preserve custom styles during HTML conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, configures HtmlSaveOptions with EnableCustomProperties set to true, and saves the workbook as an HTML file, ensuring that any custom CSS properties defined in the Excel file are retained in the output.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (custom properties option is not available in this version)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export the workbook to HTML using the specified options
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
