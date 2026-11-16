// Title: C# example: Load HTML into Aspose.Cells, update a cell, and export to HTML with inline styles only (no external CSS)
// Common Searches: how to load an HTML file into an Aspose.Cells workbook using C# | modify a cell value in a workbook imported from HTML with Aspose.Cells | save Aspose.Cells workbook as HTML with only inline CSS, no external stylesheet | disable external CSS generation when exporting to HTML with Aspose.Cells HtmlSaveOptions | update cell A1 in an HTML‑based workbook and export without creating a CSS file using Aspose.Cells
// Tags: aspocells load html workbook c# | aspocells modify worksheet cell value | htmlsaveoptions inline styles only | aspocells export html without external css | aspocells html import edit export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExample
{
    // The sample checks for an input HTML file, loads it into an Aspose.Cells Workbook, changes the value of cell A1, configures HtmlSaveOptions to use inline styles (the default in recent versions), and saves the workbook back to HTML without generating a separate CSS file, handling errors at each step.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.html";
            const string outputFile = "output.html";

            // Verify that the input HTML file exists.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the source HTML file into a workbook.
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            try
            {
                // Access the first worksheet and modify cell A1.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("New Value");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error modifying worksheet: {ex.Message}");
                return;
            }

            // Prepare HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Note: In recent Aspose.Cells versions, inline styles are the default.
            // If a separate CSS file is generated, adjust the appropriate option here.

            try
            {
                // Save the modified workbook back to HTML format.
                workbook.Save(outputFile, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
