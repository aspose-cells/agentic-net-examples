// Title: Load a TSV file into an Aspose.Cells workbook and save it as an HTML page using HtmlSaveOptions in C#
// Common Searches: c# load tsv file into Aspose.Cells workbook and export to html | asp.net convert tsv spreadsheet to html with Aspose.Cells HtmlSaveOptions | save workbook as html without embedding images using Aspose.Cells | aspose.cells html export external css not supported tsv conversion | example code for converting tsv to html with Aspose.Cells .NET
// Tags: load tsv workbook with LoadOptions | save workbook as html using HtmlSaveOptions | export active worksheet only Aspose.Cells | disable base64 image export Aspose.Cells | external css limitation Aspose.Cells html export

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of an input.tsv file, loads it into an Aspose.Cells Workbook using LoadOptions set to LoadFormat.Tsv, configures HtmlSaveOptions to export only the active worksheet and to prevent images from being embedded as Base64, notes that external CSS files are not supported in this version, and then saves the result as output.html while handling any runtime exceptions.
class TsvToHtmlConverter
{
    static void Main()
    {
        // Path to the input TSV file
        string tsvPath = "input.tsv";

        // Verify that the TSV file exists before attempting to load it
        if (!File.Exists(tsvPath))
        {
            Console.WriteLine($"Error: The file \"{tsvPath}\" was not found.");
            return;
        }

        try
        {
            // Load the TSV file into a Workbook using LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export only the active worksheet (adjust as needed)
                ExportActiveWorksheetOnly = true,

                // Do not embed images as Base64 strings
                ExportImagesAsBase64 = false

                // Note: Exporting CSS to an external file is not supported in this version of Aspose.Cells.
                // The generated HTML will contain inline CSS.
            };

            // Path for the output HTML file
            string htmlPath = "output.html";

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Conversion completed. HTML saved to \"{htmlPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
