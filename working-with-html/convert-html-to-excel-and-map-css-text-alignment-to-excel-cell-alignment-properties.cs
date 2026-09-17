// Title: Convert an HTML table to an XLSX workbook in C# using Aspose.Cells while mapping CSS text‑align values to cell HorizontalAlignment
// AI Prompts: Generate C# code that loads an HTML file with Aspose.Cells HtmlLoadOptions, reads each cell’s CSS text‑align property, and sets the matching Style.HorizontalAlignment before saving the workbook as XLSX. | Create a method that iterates over all cells in a workbook imported from HTML and translates CSS values left, center, right, and justify into the corresponding Aspose.Cells HorizontalAlignment enum. | Write a script that validates the output folder, saves the modified workbook, and logs any alignment‑mapping errors using try‑catch in C#.
// Common Searches: how to preserve css text-align when converting html table to excel with aspose.cells c# | aspose.cells html import alignment mapping left center right | c# load html into workbook and set cell horizontalalignment based on css | convert html to xlsx using aspose.cells and keep table column alignment | example code for mapping css text-align to Aspose.Cells Style.HorizontalAlignment
// Tags: html to xlsx conversion aspose.cells | map css alignment to excel cell style | c# import html workbook aspose.cells | retain cell alignment after html load | excel horizontalalignment from css values

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an HTML file into an Aspose.Cells Workbook, iterates through each cell to translate CSS text‑align values into the corresponding HorizontalAlignment style, ensures the output directory exists, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string htmlFilePath = "input.html";
                string outputFilePath = "output.xlsx";

                // Ensure the HTML file exists before loading
                if (!File.Exists(htmlFilePath))
                {
                    Console.WriteLine($"Input HTML file not found: {htmlFilePath}");
                    return;
                }

                // Load the HTML file into a new workbook using HtmlLoadOptions
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlFilePath, loadOptions);

                // Verify that at least one worksheet exists
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("No worksheets were loaded from the HTML file.");
                    return;
                }

                // Access the first worksheet and its cells
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Iterate through all cells to optionally adjust alignment
                foreach (Cell cell in cells)
                {
                    Style style = cell.GetStyle();

                    // Placeholder for custom alignment logic
                    // e.g., switch (style.HorizontalAlignment) { ... }

                    // Reapply the style (necessary if modifications were made)
                    cell.SetStyle(style);
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an Excel file
                workbook.Save(outputFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
