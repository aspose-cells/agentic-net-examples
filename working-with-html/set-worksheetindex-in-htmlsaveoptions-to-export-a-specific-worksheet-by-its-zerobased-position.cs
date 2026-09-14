// Title: Export a single worksheet to HTML by setting the active sheet index with HtmlSaveOptions in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, verifies a zero‑based worksheet index, sets Workbook.Worksheets.ActiveSheetIndex to that index, and saves only the selected sheet as HTML using Aspose.Cells HtmlSaveOptions with ExportActiveWorksheetOnly turned on. | Show how to handle out‑of‑range worksheet indexes, activate the desired sheet, configure HtmlSaveOptions for HTML output of the active worksheet only, and catch exceptions in a .NET console program.
// Common Searches: how to export only the second worksheet to HTML using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions export active worksheet only example | C# set active sheet index before saving workbook as HTML | save a specific Excel sheet as HTML with Aspose.Cells .NET | validate worksheet index range Aspose.Cells before HTML export
// Tags: HtmlSaveOptions ExportActiveWorksheetOnly Aspose.Cells | set active worksheet index C# Aspose.Cells | export single worksheet to HTML .NET | worksheet index validation Aspose.Cells | save workbook as HTML specific sheet

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, checks that a given zero‑based worksheet index is within range, makes that sheet the active one, configures HtmlSaveOptions to export only the active worksheet, and saves the result as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set the worksheet to export (zero‑based index)
            int sheetToExport = 1; // second worksheet
            if (sheetToExport < 0 || sheetToExport >= workbook.Worksheets.Count)
            {
                Console.WriteLine($"Error: Worksheet index {sheetToExport} is out of range.");
                return;
            }
            workbook.Worksheets.ActiveSheetIndex = sheetToExport;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true
            };

            // Save the selected worksheet as an HTML file
            workbook.Save(outputPath, options);
            Console.WriteLine($"Worksheet exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
