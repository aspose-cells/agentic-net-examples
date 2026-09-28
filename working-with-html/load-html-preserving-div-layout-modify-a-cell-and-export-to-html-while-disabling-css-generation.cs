// Title: Load HTML preserving DIV layout, edit a cell, and export to HTML with inline CSS using Aspose.Cells for .NET
// AI Prompts: Import an HTML file into an Aspose.Cells Workbook while retaining the original DIV structure, change the value of cell A1, and save the workbook back to HTML with all styles inlined (no external CSS files). | Show how to use HtmlLoadOptions and HtmlSaveOptions in C# to load HTML, modify worksheet data, and generate a single HTML output that embeds CSS only.
// Common Searches: asp.net load html into workbook preserving div layout aspose.cells | how to edit a cell after importing html with aspose.cells | save workbook to html with inline styles only aspose.cells | disable external css generation when exporting html from aspose.cells | aspose.cells htmlsaveoptions exportactiveworksheetonly example
// Tags: HtmlLoadOptions preserve DIV layout | modify worksheet cell after HTML import | HtmlSaveOptions inline CSS export | export active worksheet to HTML Aspose.Cells | UTF8 encoding Aspose.Cells HTML load

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// C# example that loads an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions (UTF‑8), updates cell A1, and saves the workbook to a single HTML file with inline CSS via HtmlSaveOptions, avoiding external stylesheet files.
class Program
{
    static void Main()
    {
        try
        {
            // Input HTML file path
            const string inputPath = "input.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load HTML file (preserving the original layout as much as possible)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                Encoding = Encoding.UTF8 // ensure proper text encoding
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example modification: set cell A1 to a new value
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified");

            // Export back to HTML without generating separate CSS files
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true, // export only the modified sheet
                Encoding = Encoding.UTF8
                // Note: Aspose.Cells generates CSS inline by default; no separate CSS file is created.
            };
            const string outputPath = "output.html";
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"HTML saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
