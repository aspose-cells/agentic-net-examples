// Title: C# – Load Excel with LightCells API, keep defined names that contain “Total”, and export to PDF using Aspose.Cells
// Description: Demonstrates how to create LoadOptions with a LoadFilter set to LoadDataFilterOptions.DefinedNames, load only the defined‑name section of an XLSX file, remove every name that does not include the word “Total” (case‑insensitive), and then save the trimmed workbook as a PDF. Includes file‑existence validation and proper disposal of resources.
// Keywords: Aspose.Cells | C# | LoadFilter | DefinedNames | LightCells API | filter named ranges | Excel to PDF conversion | LoadOptions | large workbook performance | named range extraction
// Common Searches: Aspose.Cells load only defined names C# | filter named ranges containing Total Aspose | LightCells API example for defined names | save filtered workbook as PDF Aspose.Cells | remove unwanted named ranges before PDF export
// Developer Intent: Load an Excel file with only its defined names, retain those that include "Total", and generate a PDF from the filtered workbook.
// Use Cases: Speed up processing of massive workbooks by loading just the defined‑name metadata. | Create financial PDFs that show only total‑related named ranges, omitting irrelevant data. | Automate cleanup of naming conventions before publishing Excel reports as PDFs.
// AI Prompts: Write C# code that also preserves defined names starting with "Sum_" while filtering for "Total". | Explain how to modify LoadOptions to load both defined names and worksheet data with LightCells API. | Provide a step‑by‑step guide to log each name removed during the filtering loop.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    // Demonstrates how to create LoadOptions with a LoadFilter set to LoadDataFilterOptions.DefinedNames, load only the defined‑name section of an XLSX file, remove every name that does not include the word “Total” (case‑insensitive), and then save the trimmed workbook as a PDF. Includes file‑existence validation and proper disposal of resources.
    class LoadFilterDefinedNamesDemo
    {
        static void Main()
        {
            // Input Excel file path (replace with your actual file)
            string inputPath = "input.xlsx";

            // Output PDF file path
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;

            try
            {
                // Create LoadOptions and configure it to load only defined names
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    // Load only defined names using LoadFilter
                    LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames)
                };

                // Load the workbook with the specified LoadOptions
                workbook = new Workbook(inputPath, loadOptions);

                // Filter defined names to keep only those containing "Total"
                // Defined names are accessed via workbook.Worksheets.Names
                NameCollection names = workbook.Worksheets.Names;

                // Iterate backwards to safely remove items while iterating
                for (int i = names.Count - 1; i >= 0; i--)
                {
                    // The Name object's identifier is accessed via the Text property
                    if (!names[i].Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                    {
                        names.RemoveAt(i);
                    }
                }

                // Save the filtered workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine("Workbook loaded, filtered, and saved to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                workbook?.Dispose();
            }
        }
    }
}
