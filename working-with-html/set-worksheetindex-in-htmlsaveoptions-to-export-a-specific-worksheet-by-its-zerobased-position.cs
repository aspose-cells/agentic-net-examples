// Title: How to export a specific worksheet to HTML using HtmlSaveOptions.WorksheetIndex in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, sets HtmlSaveOptions.WorksheetIndex to 1, and saves only that worksheet as an HTML file. | Refactor the given program to replace ExportActiveWorksheetOnly with HtmlSaveOptions.WorksheetIndex so the second sheet is exported to HTML. | Provide a C# console example that exports the third worksheet (index 2) to HTML using Aspose.Cells HtmlSaveOptions, including checks for sheet existence.
// Common Searches: Aspose.Cells export only one worksheet to HTML by index | C# HtmlSaveOptions WorksheetIndex property example | Save a specific Excel sheet as HTML using Aspose.Cells .NET | How to export the second worksheet to HTML with Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly vs WorksheetIndex comparison
// Tags: Aspose.Cells HtmlSaveOptions WorksheetIndex | export specific worksheet to HTML .NET | selective sheet export Aspose.Cells | zero‑based worksheet index HTML conversion | C# Aspose.Cells HTML save options

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, verifies the file exists, and uses HtmlSaveOptions.WorksheetIndex to specify which zero‑based worksheet should be saved as HTML. It configures the save options, writes the selected sheet to an HTML file, and includes error handling for missing files or insufficient worksheets.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export only the active worksheet
                ExportActiveWorksheetOnly = true
            };

            // Activate the second worksheet (zero‑based index)
            if (workbook.Worksheets.Count > 1)
            {
                workbook.Worksheets.ActiveSheetIndex = 1;
            }
            else
            {
                Console.WriteLine("The workbook does not contain a second worksheet.");
                return;
            }

            // Save the selected worksheet as an HTML file
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Worksheet exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
