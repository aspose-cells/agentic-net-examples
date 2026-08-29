// Title: Generate a single HTML page with embedded Base64 images from a merged Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a merged .xlsx workbook, configures HtmlSaveOptions to include all worksheets, enable presentation mode, embed images as Base64, and saves the result as one HTML file using Aspose.Cells. | Show how to verify a workbook file exists, set HtmlSaveOptions.ShowAllSheets, SaveAsSingleFile, PresentationPreference, and ExportImagesAsBase64, then export the workbook to HTML in C#.
// Common Searches: C# Aspose.Cells export merged workbook to single HTML with embedded images | How to save all sheets of an Excel file as one HTML page using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions ShowAllSheets and ExportImagesAsBase64 example | Convert merged.xlsx to HTML with base64‑encoded pictures in C#
// Tags: Aspose.Cells HtmlSaveOptions ShowAllSheets | export Excel workbook to single HTML file C# | embed Excel images as Base64 Aspose.Cells | presentation‑friendly HTML export Aspose.Cells | merged workbook HTML conversion .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The program checks for the presence of 'merged.xlsx', loads it with Aspose.Cells, configures HtmlSaveOptions to show all sheets, save as a single file, enable presentation mode, and embed images as Base64, then saves the output as 'merged_workbook.html' and prints the full path.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string workbookPath = "merged.xlsx";

                // Verify that the workbook file exists before attempting to load it
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file not found at '{Path.GetFullPath(workbookPath)}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export all worksheets into a single HTML file
                    ShowAllSheets = true,
                    SaveAsSingleFile = true,

                    // Make the HTML more presentation‑friendly
                    PresentationPreference = true,

                    // Embed images as Base64 to keep everything in one file
                    ExportImagesAsBase64 = true
                };

                // Output HTML file path
                string htmlPath = "merged_workbook.html";

                // Save the workbook as HTML
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"Workbook has been exported to HTML: {Path.GetFullPath(htmlPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
