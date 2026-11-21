// Title: Set Aspose.Cells workbook to Automatic calculation mode, recalculate all formulas, and read updated cell values in C#
// Common Searches: Aspose.Cells set calculation mode to Automatic and recalculate workbook in C# | C# code to force formula recalculation and retrieve updated cell values using Aspose.Cells | How to read cell values after recalculating formulas with Aspose.Cells .NET | Programmatically verify Excel formula results after calculation with Aspose.Cells | Change workbook calculation settings and recalculate all formulas using Aspose.Cells for .NET
// Tags: automatic calculation mode Aspose.Cells .NET | recalculate workbook formulas Aspose.Cells | access cell values after calculation Aspose.Cells | set workbook calculation mode C# | verify formula results Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file into an Aspose.Cells Workbook, ensures the calculation mode is Automatic, forces a full formula recalculation with CalculateFormula(), reads the updated values of specific cells (e.g., A1 and B2), outputs them to the console, and saves the workbook to a new file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputFile);

            // Recalculate all formulas in the workbook (default mode is Automatic)
            workbook.CalculateFormula();

            // Access cells to verify updated values
            Worksheet sheet = workbook.Worksheets[0];
            Cell cellA1 = sheet.Cells["A1"];
            Cell cellB2 = sheet.Cells["B2"];

            Console.WriteLine($"A1 = {cellA1.Value}");
            Console.WriteLine($"B2 = {cellB2.Value}");

            // Save the workbook after recalculation
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
