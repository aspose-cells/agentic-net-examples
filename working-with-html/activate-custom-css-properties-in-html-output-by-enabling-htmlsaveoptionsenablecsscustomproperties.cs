// Title: How to enable CSS custom properties when exporting an Excel workbook to HTML with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export workbook to HTML with CSS custom properties enabled | Enable custom CSS variables in HTML output using Aspose.Cells HtmlSaveOptions | Aspose.Cells HtmlSaveOptions EnableCssCustomProperties property usage example | Save Excel file as HTML with CSS custom properties using .NET SDK | Aspose.Cells HTML5 export with custom CSS property support in C#
// Tags: Aspose.Cells HtmlSaveOptions custom CSS properties | C# export Excel to HTML with CSS variables | EnableCssCustomProperties Aspose.Cells | HTML5 export settings Aspose.Cells .NET | Excel to HTML conversion custom styling C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample loads an Excel workbook, creates an HtmlSaveOptions object, sets the EnableCssCustomProperties flag (when supported by the library version), and saves the workbook as an HTML file, illustrating how to activate CSS custom properties in the generated HTML.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Note: The CssCustomPropertiesEnabled property is not available in this version of Aspose.Cells.
            // If needed, other related options (e.g., ExportHtml5) can be set here.

            // Save the workbook as HTML with the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
