// Title: Export an Excel workbook to HTML with document and custom properties as meta tags using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as an .html file, embedding the workbook’s built‑in and custom properties as <meta> elements. | Show how to enable property export in HtmlSaveOptions so that the generated HTML head contains the workbook’s metadata. | Add robust file‑existence validation and exception handling around the Excel‑to‑HTML conversion using Aspose.Cells.
// Common Searches: Aspose.Cells export Excel to HTML with meta tags for document properties C# | Include custom workbook properties in HTML output using Aspose.Cells .NET | HtmlSaveOptions ExportDocumentProperties example code | Convert .xlsx to .html preserving author and title metadata Aspose.Cells | Validate input file before saving workbook as HTML with Aspose.Cells
// Tags: Aspose.Cells HTML export with metadata | C# workbook to HTML conversion | embed document properties as meta tags | custom properties inclusion Aspose.Cells | pre‑conversion file existence check

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the presence of input.xlsx, loads it with Aspose.Cells, configures HtmlSaveOptions to export both built‑in and custom workbook properties as <meta> tags in the HTML head, and saves the result to output.html while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export built‑in document properties (author, title, etc.) as <meta> tags
                ExportDocumentProperties = true
                // Note: HtmlSaveOptions does not expose a separate ExportCustomProperties property.
                // Custom properties are included when ExportDocumentProperties is true.
            };

            // Optional: specify a folder for external resources (styles, images)
            // htmlOptions.ExportImagesAsBase64 = false;
            // htmlOptions.ImageFolder = "HtmlResources";

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
