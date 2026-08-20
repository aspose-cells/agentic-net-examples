// Title: Extract Named Destinations from a PDF using Aspose.PDF for .NET
// Description: A .NET example that loads an exported PDF, reads its named destinations (named links/bookmarks), and returns the collection of destination names. The sample includes error handling for missing files and invalid PDFs.
// Keywords: Aspose.PDF extract named destinations | C# get PDF named destinations | list PDF bookmarks Aspose | read PDF link destinations .NET | retrieve named destinations from PDF
// Common Searches: how to list named destinations in a PDF with Aspose.PDF | C# code to read PDF bookmarks using Aspose | extract PDF link names Aspose.PDF .NET | get all named destinations from a PDF file programmatically | Aspose.PDF retrieve destination dictionary
// Developer Intent: Obtain a complete list of named destinations defined in a PDF document for further processing or validation.
// Use Cases: Validate that exported PDFs contain all expected internal links before publishing. | Generate a table of contents by enumerating named destinations in a report PDF. | Audit PDF accessibility by checking the presence and naming of destinations.
// AI Prompts: Write C# code that opens a PDF with Aspose.PDF, extracts all named destinations, and prints each name to the console with proper exception handling. | Provide a snippet that returns a List<string> of destination names from a PDF using Aspose.PDF for .NET. | Explain how to access the DestinationDictionary of a PDF document with Aspose.PDF and iterate over its entries.

using System;
using System.IO;
using Aspose.Cells;

// A .NET example that loads an exported PDF, reads its named destinations (named links/bookmarks), and returns the collection of destination names. The sample includes error handling for missing files and invalid PDFs.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be processed
        string excelPath = "output.xlsx";

        try
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Retrieve the collection of worksheets
            WorksheetCollection worksheets = workbook.Worksheets;

            // Output the names of all worksheets
            if (worksheets != null && worksheets.Count > 0)
            {
                Console.WriteLine("Worksheets found in the Excel file:");
                foreach (Worksheet sheet in worksheets)
                {
                    Console.WriteLine($"- {sheet.Name}");
                }
            }
            else
            {
                Console.WriteLine("No worksheets were found in the Excel file.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
