// Title: Programmatically calculate all formulas in an Excel workbook and retrieve the results using Aspose.Cells for .NET (C#)
// Common Searches: C# code to evaluate all formulas in an Excel file with Aspose.Cells | How to get calculated values from cells after Workbook.CalculateFormula in .NET | Aspose.Cells example for loading, calculating, and saving an Excel workbook | Retrieve numeric and string cell values after formula evaluation using Aspose.Cells C# | Calculate workbook formulas and export the updated workbook with Aspose.Cells for .NET
// Tags: calculate workbook formulas Aspose.Cells C# | read evaluated cell values Aspose.Cells .NET | save workbook with calculated results Aspose.Cells | load Excel file using LoadOptions Aspose.Cells | Workbook.CalculateFormula method usage | access cell values after formula evaluation C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file with Aspose.Cells, runs Workbook.CalculateFormula() to evaluate every formula, accesses specific cells (A1 as string and B2 as numeric) to read the calculated results, prints them, and then saves the workbook containing the updated values to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied). If the file is password‑protected,
            // a CellsException will be thrown and handled in the catch block.
            var loadOptions = new LoadOptions(); // default options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve and display value from cell A1
            Cell cellA1 = sheet.Cells["A1"];
            Console.WriteLine("A1 = " + cellA1.StringValue);

            // Retrieve and display value from cell B2 (numeric example)
            Cell cellB2 = sheet.Cells["B2"];
            Console.WriteLine("B2 = " + cellB2.DoubleValue);

            // Save the workbook with calculated values
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles errors related to loading, calculating, or saving the workbook
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
