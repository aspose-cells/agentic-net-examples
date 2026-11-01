// Title: Export each Excel worksheet to its own HTML file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export individual worksheet to separate HTML page | How to save each sheet of an Excel workbook as a separate HTML file using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly example in .NET | Generate safe filenames from worksheet names when exporting to HTML with Aspose.Cells | Iterate through worksheets and export them to HTML using Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | export worksheet to html c# | save each worksheet as separate html file | sanitize worksheet name for filename c# | iterate workbook worksheets Aspose.Cells

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, checks that the file exists, then loops through every worksheet. For each sheet it makes the sheet active, configures HtmlSaveOptions with ExportActiveWorksheetOnly enabled, creates a file‑system‑safe name from the worksheet title, and saves the sheet as an individual .html file. It logs each export and catches any runtime exceptions.
class ExportWorksheetsToHtml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the current worksheet as the active one
                workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                // Configure HTML save options to export only the active worksheet
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = true
                    // ExportChartImageFormat property is not available in this version of Aspose.Cells,
                    // charts will be exported using the default image format.
                };

                // Build a safe file name based on the worksheet name
                string safeName = MakeSafeFileName(sheet.Name);
                string outputFile = $"{safeName}.html";

                // Save the active worksheet as an individual HTML file
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Worksheet '{sheet.Name}' exported to '{outputFile}'.");
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to replace invalid filename characters
    static string MakeSafeFileName(string name)
    {
        foreach (char invalidChar in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(invalidChar, '_');
        }
        return name;
    }
}
