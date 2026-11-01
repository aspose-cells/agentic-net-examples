// Title: Replace a named range with a new address in all formulas and recalculate the workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace named range reference in formulas and recalculate workbook | how to update a named range address across an entire Excel file with Aspose.Cells | C# code to change named range MyRange to Sheet1!C1:D2 in all worksheets using Aspose.Cells | recalculate formulas after modifying named ranges with Aspose.Cells .NET | iterate cells to substitute named range in formulas Aspose.Cells example
// Tags: named range reference update Aspose.Cells | replace named range in formulas C# | formula recalculation after range change Aspose.Cells | iterate cells modify formulas Aspose.Cells | load workbook modify named range save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, updates the definition of a specified named range to a new address, scans every worksheet and cell to replace that named range within formulas, recalculates all formulas, and saves the modified workbook, with comprehensive error handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            Workbook workbook = null;

            try
            {
                // Load the workbook (no password is supplied; adjust LoadOptions if a password is required)
                LoadOptions loadOptions = new LoadOptions();
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                return;
            }

            // Define the old named range and the new range reference
            string oldNamedRange = "MyRange";               // name to be replaced
            string newRangeReference = "Sheet1!C1:D2";      // new address (any valid range)

            try
            {
                // 1. Update the definition of the named range (if it exists)
                Name nameObj = workbook.Worksheets.Names[oldNamedRange];
                if (nameObj != null)
                {
                    nameObj.RefersTo = newRangeReference;
                }

                // 2. Replace occurrences of the named range inside formulas
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula) && cell.Formula.Contains(oldNamedRange))
                        {
                            cell.Formula = cell.Formula.Replace(oldNamedRange, newRangeReference);
                        }
                    }
                }

                // 3. Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // 4. Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error during processing: {ex.Message}");
            }
        }
    }
}
