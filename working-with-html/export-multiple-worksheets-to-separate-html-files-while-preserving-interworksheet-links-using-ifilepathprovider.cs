// Title: Export Multiple Worksheets to Separate HTML Files with Working Inter‑Sheet Links using Aspose.Cells IFilePathProvider (C#)
// Description: Demonstrates how to create a workbook, add cross‑sheet hyperlinks, implement a custom IFilePathProvider that returns "{SheetName}.html", configure HtmlSaveOptions to export every worksheet to its own HTML page, and preserve functional links between the generated pages. The primary output.html contains the tab strip while each sheet is saved as an individual HTML file.
// Keywords: Aspose.Cells export multiple worksheets HTML | IFilePathProvider custom file names | preserve cross‑sheet hyperlinks | HtmlSaveOptions ExportActiveWorksheetOnly false | C# workbook to separate HTML files | Aspose.Cells HTML tab strip | save workbook as individual HTML pages
// Common Searches: export each worksheet to separate html using aspose.cells | keep hyperlinks between sheets when saving as html | custom IFilePathProvider example asp.net | htmlsaveoptions export all worksheets asp.net | aspose.cells generate html files per sheet
// Developer Intent: The developer needs to save a workbook as a set of HTML files—one per worksheet—while ensuring that hyperlinks that reference other worksheets remain operational after export.
// Use Cases: Publish a multi‑page web report where Summary, Details, and Report sheets become distinct HTML pages linked together for seamless navigation. | Build a documentation site that separates content into separate HTML files per worksheet but retains cross‑references via hyperlinks. | Automate the conversion of large Excel workbooks into individually cached HTML pages to improve web‑hosting performance and SEO indexing.
// AI Prompts: Show how to modify CustomFilePathProvider to store the generated HTML files in a nested folder structure. | Explain how to embed a custom CSS stylesheet into each exported HTML sheet while still using IFilePathProvider for separate files. | Provide code that adds a navigation bar on the main output.html linking to each worksheet's individual HTML file.

using System;
using Aspose.Cells;

namespace AsposeCellsExportMultipleSheets
{
    // Provides a separate HTML file name for each worksheet based on its name.
    // Demonstrates how to create a workbook, add cross‑sheet hyperlinks, implement a custom IFilePathProvider that returns "{SheetName}.html", configure HtmlSaveOptions to export every worksheet to its own HTML page, and preserve functional links between the generated pages. The primary output.html contains the tab strip while each sheet is saved as an individual HTML file.
    internal class CustomFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Example: "Sheet1.html", "DataSheet.html", etc.
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with three worksheets.
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Name = "Summary";
                Worksheet sheet1 = workbook.Worksheets[0];

                Worksheet sheet2 = workbook.Worksheets.Add("Details");
                Worksheet sheet3 = workbook.Worksheets.Add("Report");

                // Populate some data.
                sheet1.Cells["A1"].PutValue("Welcome to the Summary sheet.");
                sheet2.Cells["A1"].PutValue("Details are listed here.");
                sheet3.Cells["A1"].PutValue("Final report content.");

                // Add an inter‑worksheet hyperlink from Summary!B2 to Details!A1.
                int link1Index = sheet1.Hyperlinks.Add(1, 1, 1, 1, "Details!A1");
                Hyperlink link1 = sheet1.Hyperlinks[link1Index];
                link1.ScreenTip = "Go to Details";
                link1.TextToDisplay = "Details Link";

                // Add another hyperlink from Details!B2 to Report!A1.
                int link2Index = sheet2.Hyperlinks.Add(1, 1, 1, 1, "Report!A1");
                Hyperlink link2 = sheet2.Hyperlinks[link2Index];
                link2.ScreenTip = "Go to Report";
                link2.TextToDisplay = "Report Link";

                // Configure HTML save options.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export each worksheet to its own HTML file.
                    ExportActiveWorksheetOnly = false,
                    // Use the custom file path provider so links point to the correct files.
                    FilePathProvider = new CustomFilePathProvider()
                };

                // Save the workbook. The main file (output.html) contains the tab strip,
                // and each worksheet is saved as a separate HTML file as defined by the provider.
                workbook.Save("output.html", saveOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
