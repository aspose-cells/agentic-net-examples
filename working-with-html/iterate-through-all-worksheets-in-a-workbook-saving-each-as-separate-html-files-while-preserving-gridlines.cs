// Title: Save each worksheet of an Excel workbook as a separate HTML file with gridlines using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, iterates through all worksheets, and writes each sheet to its own .html file while keeping the Excel gridlines visible. | Show how to set up Aspose.Cells HtmlSaveOptions to export only the active worksheet and enable ExportGridLines for HTML conversion in a .NET console application.
// Common Searches: asp.net core convert each Excel sheet to separate HTML files preserving gridlines with Aspose.Cells | c# iterate workbook worksheets and export to html using HtmlSaveOptions ExportGridLines | how to save individual worksheets as html pages while keeping gridlines in Aspose.Cells | Aspose.Cells export active worksheet only to html with gridlines enabled
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | export each worksheet to individual HTML file | C# Aspose.Cells active worksheet HTML export | preserve Excel gridlines in HTML conversion | iterate workbook worksheets Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// Loads input.xlsx, verifies the file exists, creates a Workbook, loops through all worksheets, sets each as the active sheet, configures HtmlSaveOptions with ExportGridLines=true and ExportActiveWorksheetOnly=true, saves each sheet as SheetName.html, and logs any errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the current worksheet as the active one
                workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                // Configure HTML save options to preserve gridlines
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,               // keep gridlines visible
                    ExportActiveWorksheetOnly = true      // export only the active sheet
                    // OnePagePerSheet property is not available in this version of Aspose.Cells
                };

                // Build the output file name (e.g., Sheet1.html)
                string outputFile = $"{sheet.Name}.html";

                // Save the active worksheet as an individual HTML file
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Saved: {outputFile}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
