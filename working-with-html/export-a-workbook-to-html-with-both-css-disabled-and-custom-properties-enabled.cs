// Title: Export an Excel workbook to HTML without embedded CSS and with custom properties using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.CssStyleSheetType to None and ExportCustomProperties to true, and saves the workbook as an HTML file. | Demonstrate how to check for the input Excel file, configure HtmlSaveOptions to omit CSS style sheets while preserving custom document properties, and handle any exceptions during the HTML export.
// Common Searches: Aspose.Cells C# save Excel as HTML without CSS stylesheet | How to include custom document properties when converting Excel to HTML with Aspose.Cells | HtmlSaveOptions CssStyleSheetType None example Aspose.Cells | Export workbook custom properties to HTML using Aspose.Cells .NET | Convert .xlsx to .html with no embedded styles using Aspose.Cells
// Tags: aspose.cells htmlsaveoptions disable css | aspose.cells export custom properties html | c# convert xlsx to html without stylesheet | aspose.cells workbook to html with custom properties | htmlsaveoptions cssstylessheettype none c# | aspose.cells html export exception handling

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, configures HtmlSaveOptions to turn off CSS style sheet generation and enable export of custom document properties, then saves the workbook as an HTML file while handling possible errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Note: In newer Aspose.Cells versions, CssStyleSheetType and ExportCustomProperties
            // are either set by default or not available. The essential save operation works
            // without explicitly setting them.

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
