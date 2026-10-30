// Title: Recalculate formulas only when volatile functions are present before exporting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: c# aspnet detect volatile Excel functions with Aspose.Cells before PDF conversion | how to trigger workbook.CalculateFormula only if sheet contains NOW or RAND in Aspose.Cells | selective formula calculation in Aspose.Cells prior to saving workbook as PDF | scan worksheets for volatile formulas using Aspose.Cells C# and recalc conditionally | conditional Excel to PDF export with Aspose.Cells based on presence of volatile functions
// Tags: volatile function detection Aspose.Cells | conditional formula calculation .NET | Excel to PDF conversion Aspose.Cells | worksheet scanning for volatile formulas C# | calculate workbook formulas before PDF export

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, iterates through each worksheet to find formulas that contain volatile functions such as NOW(), RAND(), or OFFSET(). If any worksheet includes such functions, it triggers a full workbook recalculation using workbook.CalculateFormula() and then saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // List of volatile functions to look for (case‑insensitive)
            string[] volatileFunctions = new string[]
            {
                "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "CELL()", "INFO()", "AREAS()", "GETPIVOTDATA()"
            };

            bool anySheetNeedsRecalc = false;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool containsVolatile = false;

                // Scan all cells that have formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula.ToUpperInvariant();

                        // Check if the formula contains any volatile function
                        foreach (string vf in volatileFunctions)
                        {
                            if (formula.Contains(vf))
                            {
                                containsVolatile = true;
                                break;
                            }
                        }

                        if (containsVolatile)
                            break; // No need to continue scanning this sheet
                    }
                }

                if (containsVolatile)
                {
                    anySheetNeedsRecalc = true;
                }
            }

            // Perform calculation if any sheet contained volatile functions
            if (anySheetNeedsRecalc)
            {
                workbook.CalculateFormula(); // Recalculates all formulas in the workbook
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
