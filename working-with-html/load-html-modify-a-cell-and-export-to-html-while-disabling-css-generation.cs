// Title: C# example: Load HTML into Aspose.Cells, update cell A1, and save as HTML while preventing CSS output
// AI Prompts: Write C# code that uses Aspose.Cells to open an HTML file, set the value of cell A1 in the first worksheet, and save the workbook to a new HTML file with CSS generation turned off. | Show how to configure Aspose.Cells HtmlSaveOptions in .NET to suppress stylesheet creation when exporting a modified HTML workbook.
// Common Searches: aspnet aspose.cells load html file modify cell and export without css | how to prevent css stylesheet when saving workbook to html using Aspose.Cells C# | change value of A1 in html workbook using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions disable css generation example
// Tags: Aspose.Cells load HTML workbook C# | modify cell value Aspose.Cells | HtmlSaveOptions disable CSS Aspose.Cells | export workbook to HTML without stylesheet | Aspose.Cells HTML import export example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExample
{
    // // Loads an HTML file into an Aspose.Cells Workbook, updates cell A1, and saves the workbook back to HTML. Current Aspose.Cells versions do not expose a direct ExportCss flag, so the example uses default HtmlSaveOptions while noting that CSS suppression is not available.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputFile = "input.html";
                const string outputFile = "output.html";

                // Verify that the input HTML file exists to avoid FileNotFoundException.
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: The input file \"{inputFile}\" was not found.");
                    return;
                }

                // Load the source HTML file into a workbook.
                Workbook workbook = new Workbook(inputFile);

                // Modify a specific cell (e.g., A1) in the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("New Value");

                // Prepare HTML save options.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                // Note: In recent Aspose.Cells versions the ExportCss property is not available.
                // If CSS generation needs to be disabled, configure the appropriate options here
                // when such a property exists. For now, default behavior is used.

                // Export the modified workbook back to HTML.
                workbook.Save(outputFile, saveOptions);
                Console.WriteLine($"Workbook successfully saved to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
