// Title: Load an Excel workbook and access the second worksheet (index 1) using Aspose.Cells for .NET to enable formula verification
// AI Prompts: Open a .xlsx file with Aspose.Cells, retrieve the worksheet at zero‑based index 1, and store it for subsequent formula checks. | After loading the workbook, call Workbook.CalculateFormula() to evaluate all formulas, then use the obtained second worksheet for validation.
// Common Searches: Aspose.Cells .NET how to get the second worksheet by index | C# calculate all formulas after loading an Excel workbook with Aspose.Cells | prepare worksheet for formula validation using Aspose.Cells in a console application | access worksheet at index 1 in Aspose.Cells and run full calculation
// Tags: Aspose.Cells load workbook .xlsx C# | Aspose.Cells get worksheet by zero‑based index | Aspose.Cells calculate all formulas | prepare second worksheet for formula verification | C# Excel formula validation with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing .xlsx file, verifies its existence, opens it with Aspose.Cells, accesses the second worksheet (index 1), forces a full formula calculation, and leaves the worksheet ready for further formula verification.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the second worksheet (index 1 because the collection is zero‑based)
            Worksheet secondWorksheet = workbook.Worksheets[1];

            // Force an immediate calculation of all formulas
            workbook.CalculateFormula();

            // The secondWorksheet variable is now ready for further formula verification steps
            Console.WriteLine("Workbook loaded and formulas calculated successfully.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
