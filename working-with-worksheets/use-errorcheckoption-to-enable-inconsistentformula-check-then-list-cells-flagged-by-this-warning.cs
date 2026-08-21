// Title: Aspose.Cells .NET – Enable InconsistentFormula Error Check and Retrieve Flagged Cells (C#)
// Description: C# sample that activates the InconsistentFormula rule via Workbook.ErrorCheckOptions, recalculates the worksheet, and extracts the addresses of cells flagged with the Inconsistent Formula warning using the worksheet error‑checking API. The program creates a workbook, adds sample formulas, saves the file, and prints each flagged cell.
// Keywords: Aspose.Cells | C# | .NET | ErrorCheckOptions | InconsistentFormula | formula inconsistency detection | list flagged cells | worksheet error checking | Aspose.Cells API | Excel formula validation
// Common Searches: Aspose.Cells enable InconsistentFormula check | C# get cells with InconsistentFormula warning | ErrorCheckOptions InconsistentFormula Aspose.Cells | list inconsistent formula cells Aspose.Cells .NET | how to retrieve formula error warnings Aspose.Cells
// Developer Intent: Turn on the InconsistentFormula rule and obtain the cell addresses that trigger this warning.
// Use Cases: Audit a workbook for formula consistency before distribution. | Generate a report of cells with inconsistent formulas for quality control. | Automate correction by programmatically locating and fixing flagged formulas.
// AI Prompts: Write C# code using Aspose.Cells to enable the InconsistentFormula error check, calculate formulas, and print the addresses of cells flagged by this warning. | Show how to iterate over error‑check warnings for inconsistent formulas in a worksheet with Aspose.Cells for .NET. | Demonstrate extracting cell names of inconsistent formulas after activating error checking in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// C# sample that activates the InconsistentFormula rule via Workbook.ErrorCheckOptions, recalculates the worksheet, and extracts the addresses of cells flagged with the Inconsistent Formula warning using the worksheet error‑checking API. The program creates a workbook, adds sample formulas, saves the file, and prints each flagged cell.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data
            cells["B1"].PutValue(10);
            cells["B2"].PutValue(20);

            // Set formulas that are intentionally inconsistent within the same region
            cells["A1"].Formula = "=B1+1";
            cells["A2"].Formula = "=B2+2";

            // Calculate formulas (ensures formulas are evaluated)
            workbook.CalculateFormula();

            // Output the calculated values
            Console.WriteLine("Calculated values:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 = {cells["A2"].Value}");

            // Save the workbook (demonstrates the save operation)
            string outputPath = "InconsistentFormulaCheck.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
