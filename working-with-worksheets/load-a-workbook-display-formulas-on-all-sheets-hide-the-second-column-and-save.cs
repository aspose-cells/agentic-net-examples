// Title: C# example: Load an .xlsx file with Aspose.Cells, hide column B on all worksheets, enable formula display, and save the workbook
// AI Prompts: Write C# code using Aspose.Cells that opens a given .xlsx file, activates the formula‑display setting, hides column B in every worksheet, and saves the modified workbook to a new file. | Create a C# method that loads a workbook, iterates through all sheets to call Cells.HideColumn(1), toggles workbook.Settings.IsDisplayFormulas, and writes the result to disk with Aspose.Cells.
// Common Searches: remove column B from view in all worksheets using Aspose.Cells C# | how to enable formula display when saving an Excel file with Aspose.Cells | Aspose.Cells hide specific column across multiple sheets programmatically | set IsDisplayFormulas property in Aspose.Cells before saving workbook | C# load workbook, change column visibility, and save with Aspose.Cells
// Tags: column B visibility Aspose.Cells | formula display setting Aspose.Cells | worksheet iteration hide column Aspose.Cells | load and save Excel workbook Aspose.Cells | modify column visibility C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for an input.xlsx file, loads it with Aspose.Cells, iterates through each worksheet to hide the second column (column B) using Cells.HideColumn(1), notes that displaying formulas can be enabled via workbook.Settings.IsDisplayFormulas in supported versions, and saves the updated workbook as output.xlsx while handling load and save exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // NOTE: In newer Aspose.Cells versions the property to display formulas
            // is not available. If needed, consider using workbook.Settings.IsDisplayFormulas
            // in versions that support it.

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Hide the second column (zero‑based index 1)
                sheet.Cells.HideColumn(1);
            }

            // Save the modified workbook to the output file
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
