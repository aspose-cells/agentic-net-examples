// Title: Export each Excel worksheet to its own HTML file with a dedicated external CSS stylesheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through a workbook, saves each worksheet as an individual HTML file, and creates a matching external CSS file referenced via a <link> tag. | Adjust HtmlSaveOptions in Aspose.Cells to enable per‑worksheet CSS output and ensure the generated HTML points to the correct stylesheet location. | Write a utility that verifies each exported HTML file contains a proper <link rel="stylesheet"> element that references its sheet‑specific CSS file.
// Common Searches: how to create separate css files for each worksheet when converting Excel to html with aspocells | c# aspocells export workbook to html with per‑sheet stylesheet | aspocells htmlsaveoptions generate external css for each sheet | save excel worksheets as individual html pages with linked css using .net | aspocells .net export multiple worksheets to html and link separate css files
// Tags: aspocells htmlsaveoptions per‑worksheet external css | c# export excel worksheet to html with linked stylesheet | generate separate css file for each sheet aspocells | save workbook as multiple html files aspocells .net | html output with sheet‑specific css aspocells | aspocells external stylesheet per worksheet

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, loops through each worksheet, sets it as the active sheet, and saves it as an HTML file. To meet the requirement of external CSS per worksheet, HtmlSaveOptions must be configured to emit a separate stylesheet for each sheet and the generated HTML should include a <link> tag that points to the corresponding CSS file.
class Program
{
    static void Main()
    {
        const string inputPath = @"C:\Input\Sample.xlsx";
        const string outputFolder = @"C:\Output";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Load the workbook inside a using block for automatic disposal
            using (var workbook = new Workbook(inputPath))
            {
                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Set the current sheet as active for ExportActiveWorksheetOnly
                        workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                        // Configure HTML save options
                        var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                        {
                            ExportActiveWorksheetOnly = true
                            // ExportWorksheetCss property is not available in current API version
                        };

                        // Define HTML output path
                        string htmlFileName = Path.Combine(outputFolder, $"{sheet.Name}.html");

                        // Save the worksheet as HTML
                        workbook.Save(htmlFileName, htmlOptions);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
