// Title: Export an Excel workbook to HTML with correct rendering of multi‑row merged cells using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file, configures HtmlSaveOptions (ExportGridLines = true, Encoding = UTF8), and saves it as an .html file while preserving any merged cells that span multiple rows. | Show how to programmatically merge a range of cells in a worksheet before exporting to HTML with Aspose.Cells, ensuring the merged region appears correctly in the output. | Provide robust error‑handling for exporting an Excel workbook to HTML with Aspose.Cells, including file‑existence verification and exception logging.
// Common Searches: how to keep merged cells when converting Excel to HTML using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions preserve row spanning merged regions | C# export .xlsx to .html with grid lines and UTF‑8 encoding | export Excel workbook to HTML while maintaining cell merges Aspose | sample code for saving Excel as HTML with merged cells Aspose.Cells
// Tags: Aspose.Cells HTML export with merged cells | C# HtmlSaveOptions ExportGridLines UTF-8 | preserve row span merged region Aspose.Cells | Excel to HTML conversion .NET Aspose.Cells | error handling file existence Aspose.Cells export

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example checks for the input.xlsx file, loads it with Aspose.Cells, optionally merges a range of cells, sets HtmlSaveOptions to export grid lines and use UTF‑8 encoding, and saves the workbook as output.html while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the source Excel file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Example of creating a merged region (optional)
            // workbook.Worksheets[0].Cells.Merge(1, 0, 3, 2); // Merges rows 1‑3 and columns A‑C

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Preserve grid lines for visual fidelity (optional)
                ExportGridLines = true,

                // Use UTF‑8 encoding for the output HTML
                Encoding = Encoding.UTF8
            };

            // Export the workbook to an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
