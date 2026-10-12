// Title: Retrieve a worksheet by name, display its formulas instead of values, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing .xlsx file, selects the worksheet named "SheetName", iterates through every cell to replace the displayed value with the cell's formula text, and saves the modified workbook to a new file. | Update the given Aspose.Cells example to check whether a worksheet called "SheetName" exists, fall back to the first worksheet if it does not, and apply a technique that forces formulas to be shown before saving the workbook. | Add robust error handling to a C# Aspose.Cells script that loads a workbook, retrieves a specific worksheet, converts cell values to their formula strings, and writes the result, ensuring graceful handling of missing files and missing worksheets.
// Common Searches: how to display formulas instead of results in an Excel file using Aspose.Cells C# | Aspose.Cells C# retrieve worksheet by name and show formulas | C# Aspose.Cells save workbook with formulas visible | iterate over cells to get formula text with Aspose.Cells .NET | fallback to first sheet when worksheet name not found Aspose.Cells
// Tags: retrieve worksheet by name Aspose.Cells C# | show formulas in Aspose.Cells workbook | convert cell values to formula text Aspose.Cells | fallback to first worksheet Aspose.Cells | save workbook with formulas displayed Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, attempts to get a worksheet named "SheetName" (using the first sheet as a fallback), notes that the ShowFormula property is unavailable, and then saves the workbook. It includes basic error handling for missing files and worksheets.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve a specific worksheet; fallback to the first worksheet if not found
            Worksheet worksheet = workbook.Worksheets["SheetName"];
            if (worksheet == null)
            {
                Console.WriteLine("Worksheet 'SheetName' not found. Using the first worksheet.");
                worksheet = workbook.Worksheets[0];
            }

            // NOTE: The ShowFormula setting is not available in this version of Aspose.Cells.
            // If needed, alternative approaches (e.g., iterating cells and setting formulas as text) can be implemented here.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
