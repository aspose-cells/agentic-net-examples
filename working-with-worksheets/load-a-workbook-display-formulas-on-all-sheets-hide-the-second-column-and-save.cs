// Title: Load an XLSX workbook, hide column B on every worksheet, and save the modified file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing XLSX file with Aspose.Cells, iterates through each worksheet, hides the second column, and writes the result to a new file. | Show how to use Aspose.Cells Workbook and Worksheet objects to conceal column B across all sheets and then save the workbook. | Explain how to verify the input file, hide the second column, and handle exceptions while processing an Excel workbook with Aspose.Cells in C#.
// Common Searches: how to hide column B in every worksheet using Aspose.Cells C# | Aspose.Cells hide second column across all sheets and save workbook | C# load Excel file, hide column, and export with Aspose.Cells | display formulas not supported Aspose.Cells version hide column example
// Tags: Aspose.Cells hide second column across worksheets | Aspose.Cells workbook load and save XLSX | Aspose.Cells iterate worksheets C# | Aspose.Cells formula display limitation | C# hide Excel column with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing XLSX workbook, attempts to show formulas (unsupported in the current Aspose.Cells version), hides column B (index 1) on every worksheet, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // NOTE: The ShowFormula setting is not available in the current Aspose.Cells version.
            // If needed, upgrade the library or use an alternative approach.

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Hide the second column (zero‑based index 1)
                sheet.Cells.HideColumn(1);
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
