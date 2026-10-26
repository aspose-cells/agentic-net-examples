// Title: Generate separate HTML files with external CSS for each worksheet using Aspose.Cells for .NET
// Common Searches: how to export each Excel worksheet to its own HTML file with external CSS using Aspose.Cells C# | Aspose.Cells create per‑worksheet CSS files when saving as HTML in .NET | C# save workbook worksheets as HTML with linked stylesheet files Aspose.Cells | export Excel sheets to HTML with separate CSS files and correct link paths using Aspose.Cells | Aspose.Cells HtmlSaveOptions external CSS per worksheet example C#
// Tags: Aspose.Cells per-worksheet HTML export external CSS | HtmlSaveOptions ExportActiveWorksheetOnly linked stylesheet | C# generate separate CSS files for Excel worksheets | Excel to HTML conversion external stylesheet Aspose.Cells | output directory structure worksheet HTML CSS Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates an output folder, iterates through each worksheet, activates it, and saves it as an individual HTML file. Aspose.Cells generates a dedicated external CSS file for each worksheet and updates the HTML <link> tag so the stylesheet path points to the correct location.
class ExportWorksheetsToHtmlWithExternalCss
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string workbookPath = @"C:\Input\SampleWorkbook.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Output directory for HTML files (and any generated CSS/images)
            string outputDir = @"C:\Output\HtmlExport";
            Directory.CreateDirectory(outputDir);

            // Export each worksheet to its own HTML file
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Activate the current worksheet
                workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                // Configure HTML save options – export only the active worksheet
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = true
                };

                // Build the HTML file path (one file per worksheet)
                string htmlFileName = $"{sheet.Name}.html";
                string htmlPath = Path.Combine(outputDir, htmlFileName);

                // Save the worksheet as HTML
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine($"Exported worksheet '{sheet.Name}' to HTML: {htmlPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
