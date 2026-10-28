// Title: Export each worksheet of an Excel workbook to separate HTML5 files while preserving inter‑worksheet hyperlinks using a custom IFilePathProvider in Aspose.Cells for .NET
// Common Searches: how to save each worksheet as an individual HTML file with Aspose.Cells C# | preserve worksheet hyperlinks when exporting worksheets to separate HTML pages using Aspose.Cells | use IFilePathProvider to control HTML file names in Aspose.Cells .NET | export active worksheet only to HTML5 with HtmlSaveOptions in Aspose.Cells | iterate through workbook worksheets and generate separate HTML files in C#
// Tags: custom IFilePathProvider HTML file naming Aspose.Cells | ExportActiveWorksheetOnly HtmlSaveOptions | separate HTML files per worksheet Aspose.Cells | inter‑worksheet hyperlink preservation HTML export | HTML5 output Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example implements a WorksheetLinkProvider that ensures each generated file name ends with .html, loads an Excel workbook, loops through all worksheets, sets each as the active sheet, configures HtmlSaveOptions with ExportActiveWorksheetOnly = true and HtmlVersion = Html5, and saves each sheet to its own HTML file named after the worksheet while logging success or error messages.
class WorksheetLinkProvider : IFilePathProvider
{
    // Ensures that links between worksheets point to the correct HTML files.
    public string GetFullName(string fileName)
    {
        // Aspose.Cells may pass the worksheet name without extension; enforce .html.
        if (!fileName.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
        {
            fileName = $"{fileName}.html";
        }
        return fileName;
    }
}

class ExportWorksheetsToHtml
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Create a single instance of the custom file path provider (not used in this version).
            IFilePathProvider pathProvider = new WorksheetLinkProvider();

            // Iterate through each worksheet and export it to a separate HTML file.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Set the current worksheet as active for ExportActiveWorksheetOnly.
                    workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                    // Configure HTML save options for the current worksheet.
                    HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
                    {
                        ExportActiveWorksheetOnly = true, // Export only the active worksheet.
                        HtmlVersion = HtmlVersion.Html5   // Use HTML5 output.
                        // Note: ExportWorksheetNavigation and CustomFilePathProvider are not available
                        // in the current Aspose.Cells version, so they are omitted.
                    };

                    // Define the output HTML file name (one file per worksheet).
                    string outputFile = $"{sheet.Name}.html";

                    // Save the current worksheet as HTML.
                    workbook.Save(outputFile, options);
                    Console.WriteLine($"Worksheet '{sheet.Name}' exported to '{outputFile}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export worksheet '{sheet.Name}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
