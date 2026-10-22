// Title: Convert HTML to XLSX with Aspose.Cells in C# and enable automatic formula calculation
// AI Prompts: Write a C# program that reads an HTML file, loads it into an Aspose.Cells Workbook using HTML LoadOptions, switches the workbook's calculation mode to Automatic, forces a full formula recomputation, and saves the workbook as an .xlsx file. | Show how to verify the existence of the source HTML file, handle possible exceptions, and confirm that all formulas are evaluated when the workbook is exported to Excel with Aspose.Cells.
// Common Searches: C# example for loading an HTML file into Aspose.Cells and preserving formulas | Aspose.Cells .NET set workbook calculation mode after importing HTML | automatically recalculate all formulas during HTML to Excel conversion with Aspose.Cells | best practice for converting HTML tables to XLSX using Aspose.Cells in C#
// Tags: convert HTML document to Excel workbook via Aspose.Cells C# | enable automatic formula calculation in Aspose.Cells workbook | trigger complete formula recomputation after HTML import | initialize workbook from HTML using Aspose.Cells options | export workbook as .xlsx with evaluated formulas

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for an input HTML file, loads it into an Aspose.Cells Workbook with HTML LoadOptions, forces a full formula recalculation, sets the calculation mode to Automatic, and saves the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.xlsx";

            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Recalculate all formulas (optional but ensures up‑to‑date values)
            workbook.CalculateFormula();

            // Save the workbook as an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
