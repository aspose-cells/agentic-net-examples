// Title: Recalculate all formulas in an Excel workbook after changing cell values using Aspose.Cells Workbook.CalculateFormula in C#
// Common Searches: C# Aspose.Cells recalculate workbook formulas after updating cell values | How to force formula recalculation in an Excel file with Aspose.Cells | Using Workbook.CalculateFormula to refresh formulas after modifying worksheet data in .NET | Aspose.Cells update formulas programmatically after cell edit C# | Recompute all Excel formulas in a saved workbook with Aspose.Cells .NET
// Tags: Workbook.CalculateFormula C# | Aspose.Cells formula recalculation | programmatic Excel formula update | modify worksheet cells Aspose.Cells | recalculate all workbook formulas .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new workbook if missing), changes the values of cells A1 and B1 on the first worksheet, calls Workbook.CalculateFormula to recompute every formula in the workbook, and saves the updated file, handling password‑protected workbooks and general exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists
            if (File.Exists(inputPath))
            {
                try
                {
                    // Attempt to load; if the file is password‑protected an exception will be thrown
                    var loadOptions = new LoadOptions();
                    workbook = new Workbook(inputPath, loadOptions);
                }
                catch (CellsException ex) when (ex.Message.Contains("password"))
                {
                    Console.WriteLine("The input workbook is password protected. Unable to load it.");
                    return;
                }
            }
            else
            {
                // Create a new workbook when the input file is missing
                workbook = new Workbook();
                Console.WriteLine("Input file not found. A new workbook has been created.");
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Modify some cell values
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook with updated values and formulas
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
