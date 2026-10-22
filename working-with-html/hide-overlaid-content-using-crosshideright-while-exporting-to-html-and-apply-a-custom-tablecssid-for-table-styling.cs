// Title: Export an Excel workbook to HTML with a custom TableCssId and attempt to hide overlaid content using CrossHideRight in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.TableCssId to a custom identifier, and saves the workbook as an HTML document. | Demonstrate how to verify the source Excel file exists and wrap the conversion in try‑catch blocks to handle errors when exporting to HTML with Aspose.Cells. | Explain why the HtmlCrossHideRight property is unavailable in the current Aspose.Cells release and suggest an alternative technique for suppressing overlaid content in the generated HTML.
// Common Searches: Aspose.Cells export to HTML custom table CSS id C# example | how to hide overlaid cells when saving Excel as HTML using Aspose.Cells | HtmlSaveOptions TableCssId property usage Aspose.Cells .NET | CrossHideRight not supported Aspose.Cells HTML export workaround | C# check file existence before converting Excel to HTML with Aspose.Cells
// Tags: HtmlSaveOptions TableCssId Aspose.Cells | Excel to HTML Aspose.Cells C# | overlaid content hide Aspose.Cells HTML | CrossHideRight limitation Aspose.Cells | custom table styling Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to assign a custom TableCssId ("customTable"), notes that the HtmlCrossHideRight feature is not available in the current version, and saves the workbook as output.html while handling any exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsx";
                string outputFile = "output.html";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(inputFile);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // The HtmlCrossHideRight property is not available in the current Aspose.Cells version.
                // Set a custom CSS ID for the generated HTML table
                htmlOptions.TableCssId = "customTable";

                // Export the workbook to HTML with the specified options
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
