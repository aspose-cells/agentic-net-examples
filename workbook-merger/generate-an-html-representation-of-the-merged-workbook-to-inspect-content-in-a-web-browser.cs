// Title: Export a Merged Workbook to a Single HTML File with All Sheets – Aspose.Cells for .NET (C#)
// Description: Loads a merged Excel file (merged.xlsx) using Aspose.Cells, configures HtmlSaveOptions to combine every worksheet into one HTML document, applies presentation‑friendly styling, embeds images as Base64, saves the result as merged_workbook.html, and verifies the output file.
// Keywords: Aspose.Cells HTML export | C# save workbook as single HTML | HtmlSaveOptions ShowAllSheets | PresentationPreference Aspose.Cells | ExportImagesAsBase64 | merged Excel to HTML | .NET Excel to web preview
// Common Searches: Aspose.Cells export merged workbook to one HTML file | C# convert Excel with multiple sheets to single HTML page | How to embed images as Base64 when saving Excel as HTML | Show all worksheets in one HTML file using Aspose.Cells | Presentation‑friendly HTML output from Excel .NET
// Developer Intent: Generate a self‑contained HTML representation of a merged Excel workbook that includes every worksheet in a single file.
// Use Cases: Quickly preview combined workbook data in any browser without Excel. | Create a portable HTML report for email distribution or documentation. | Embed the HTML view in a web portal to let users inspect merged data.
// AI Prompts: Write C# code that uses Aspose.Cells to export a Workbook to one HTML file with ShowAllSheets and ExportImagesAsBase64 enabled. | Explain how HtmlSaveOptions.PresentationPreference changes the layout of the generated HTML. | Show how to modify the code to produce separate HTML files for each worksheet instead of a single file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Loads a merged Excel file (merged.xlsx) using Aspose.Cells, configures HtmlSaveOptions to combine every worksheet into one HTML document, applies presentation‑friendly styling, embeds images as Base64, saves the result as merged_workbook.html, and verifies the output file.
    class Program
    {
        static void Main()
        {
            // Load the merged workbook (replace with actual path)
            string workbookPath = "merged.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Save everything (all worksheets) into a single HTML file
            htmlOptions.SaveAsSingleFile = true;
            htmlOptions.ShowAllSheets = true;
            // Optional: make the HTML more presentation‑friendly
            htmlOptions.PresentationPreference = true;
            // Optional: embed images as Base64 to keep a single file
            htmlOptions.ExportImagesAsBase64 = true;

            // Define output HTML file path
            string htmlPath = "merged_workbook.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the file was created
            if (File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML representation saved successfully to: {htmlPath}");
            }
            else
            {
                Console.WriteLine("Failed to create the HTML file.");
            }
        }
    }
}
