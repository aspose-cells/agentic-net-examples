// Title: Load HTML into an Aspose.Cells Workbook while preserving DIV tag layout using HtmlLoadOptions (C#)
// AI Prompts: Write C# code that loads a local HTML file into an Aspose.Cells Workbook with DIV tag layout retained via HtmlLoadOptions. | Demonstrate how to set HtmlLoadOptions to enable DIV layout preservation when converting HTML to an XLSX workbook with Aspose.Cells for .NET. | Create a C# console application that checks for an HTML file, loads it into a Workbook preserving DIV elements, and saves it as an Excel file.
// Common Searches: Aspose.Cells C# load html file preserve div layout | How to keep div tags when converting HTML to Excel with Aspose.Cells | HtmlLoadOptions EnableDivTagLayout example .NET | Convert HTML to XLSX while maintaining div structure using Aspose.Cells | C# Aspose.Cells load html with div layout enabled
// Tags: Aspose.Cells HtmlLoadOptions preserve div layout | load html into workbook Aspose.Cells .NET | convert html to xlsx with div structure | EnableDivTagLayout Aspose.Cells C# | HTML to Excel conversion retaining layout

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the specified HTML file exists, loads it into an Aspose.Cells Workbook using HtmlLoadOptions (which automatically preserves DIV tag layout), and saves the workbook as an XLSX file, handling any exceptions that may occur.
class HtmlToWorkbook
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlFilePath = "input.html";

            // Verify that the HTML file exists before attempting to load it
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: The file \"{htmlFilePath}\" was not found.");
                return;
            }

            // Configure HTML load options (default options are sufficient for DIV layout preservation)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Define the output Excel file path
            string outputFilePath = "output.xlsx";

            // Save the workbook to an Excel file (XLSX format)
            workbook.Save(outputFilePath, SaveFormat.Xlsx);

            Console.WriteLine("HTML file has been loaded and saved as Excel with DIV layout preserved.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
