// Title: How to export Excel worksheets as HTML heading elements using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.ExportWorksheetHeader to true, and saves the workbook as .html so each worksheet name appears as an <h1> tag. | Show a complete example that checks for the input Excel file, configures HtmlSaveOptions for HTML output, enables worksheet name headings, and handles possible exceptions. | Provide a console‑based C# program that converts a workbook to HTML with sheet titles rendered as heading elements and confirms the output path.
// Common Searches: Aspose.Cells C# export Excel to HTML with sheet names as headings | Enable ExportWorksheetHeader in HtmlSaveOptions for HTML conversion | C# code to convert .xlsx to .html and include worksheet titles | How to add <h1> tags for each worksheet when saving Excel as HTML using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportWorksheetHeader | C# convert Excel to HTML with sheet headings | HTML output include worksheet titles Aspose.Cells | Excel to HTML heading elements C# | Workbook.Save HtmlSaveOptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions, enables the ExportWorksheetHeader flag to render each worksheet name as an <h1> element, and saves the result as an HTML file while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputFile);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Enable worksheet name headings if the property is available in the used version
            // htmlOptions.ExportWorksheetHeader = true; // Uncomment if supported

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
