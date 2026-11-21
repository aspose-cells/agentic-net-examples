// Title: Load HTML into an Aspose.Cells Workbook while disabling DIV tag layout (C#)
// Common Searches: C# Aspose.Cells load HTML file without preserving DIV tags | How to disable DIV layout when importing HTML into Aspose.Cells workbook | Aspose.Cells HtmlLoadOptions default behavior ignore DIV layout C# | Convert HTML to XLSX with Aspose.Cells while ignoring DIV elements
// Tags: HtmlLoadOptions EnableDivTagLayout false | Aspose.Cells HTML to Excel conversion | C# import HTML into workbook ignoring DIV tags | Disable DIV layout during HTML import | Load HTML file into workbook using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the input HTML file, creates a default HtmlLoadOptions instance where EnableDivTagLayout is false, loads the HTML into an Aspose.Cells Workbook, and saves the result as an XLSX file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Verify that the input HTML file exists
        if (!File.Exists(htmlFilePath))
        {
            Console.WriteLine($"Error: The file '{htmlFilePath}' was not found.");
            return;
        }

        try
        {
            // Configure HTML load options (default behavior does not preserve DIV layout)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook to an Excel file (XLSX format)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
