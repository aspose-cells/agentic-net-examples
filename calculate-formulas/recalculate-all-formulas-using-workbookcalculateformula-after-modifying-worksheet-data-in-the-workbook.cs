// Title: Recalculate all formulas in an Excel workbook after updating cell values using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate workbook formulas after changing cell values | force formula evaluation in an Excel file with Workbook.CalculateFormula | update Excel formulas programmatically before saving using Aspose.Cells | recompute all formulas in a .xlsx after modifying cells with Aspose.Cells .NET | C# example to recalc formulas after editing worksheet data with Aspose.Cells
// Tags: Workbook.CalculateFormula usage | recalculate Excel formulas Aspose.Cells | modify cell values then recalc formulas C# | Aspose.Cells formula evaluation before save | password protected workbook handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook (or creates a new one), changes the values of cells A1 and B1 on the first worksheet, calls Workbook.CalculateFormula to recompute every formula in the workbook, and then saves the updated file. It also demonstrates handling password‑protected files and proper resource disposal.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex)
                {
                    // Handle password‑protected workbook or other loading issues.
                    if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine($"The workbook \"{inputPath}\" is password protected. Unable to open it.");
                        return;
                    }
                    else
                    {
                        Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                        return;
                    }
                }
            }
            else
            {
                // No input file – create a new workbook with a default worksheet.
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Modify data: set A1 to 10 and B1 to 20.
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);

            // Recalculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Save the updated workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        finally
        {
            // Release resources if needed.
            workbook?.Dispose();
        }
    }
}
