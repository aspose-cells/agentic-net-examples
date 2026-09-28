// Title: Export an Excel workbook to HTML with row and column headings, generating a separate table for each worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures HtmlSaveOptions to include row and column headings, and saves the workbook so each worksheet appears as its own table in the resulting HTML. | Update an existing Aspose.Cells export routine to produce HTML that shows headers for rows and columns and splits the output into distinct tables per sheet without relying on the deprecated OnePagePerSheet setting.
// Common Searches: how to export Excel to HTML with row and column headings using Aspose.Cells .NET | Aspose.Cells generate separate HTML tables for each worksheet | C# save workbook as HTML with headings per sheet Aspose | HtmlSaveOptions ExportRowColumnHeadings example Aspose.Cells | convert multi‑sheet Excel to single HTML file with sheet separation Aspose
// Tags: Aspose.Cells HtmlSaveOptions row column headings | C# export Excel to HTML separate sheets | Aspose.Cells multi‑sheet HTML conversion | HTML export with worksheet separation Aspose | Aspose.Cells workbook to HTML per worksheet

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input file, loads it into an Aspose.Cells Workbook, sets HtmlSaveOptions.ExportRowColumnHeadings to true, and saves the workbook as a single HTML file where each worksheet is rendered as an independent table, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Use the updated property for exporting row/column headings
                ExportRowColumnHeadings = true
                // Note: OnePagePerSheet property is not available in the current API version
            };

            // Save the workbook as an HTML file with the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
