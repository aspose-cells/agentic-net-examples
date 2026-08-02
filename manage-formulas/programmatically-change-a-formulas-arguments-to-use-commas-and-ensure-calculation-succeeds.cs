// Title: Fix Excel formula delimiter (semicolon → comma) with Aspose.Cells for .NET and recalculate the workbook
// Description: Creates a new workbook, writes sample values, converts an invalid formula that uses a semicolon ("=SUM(A1;A2)") to a comma‑separated version, sets the corrected formula in a cell, forces Aspose.Cells to recalculate all formulas, outputs the result, and saves the file as FormulaFixed.xlsx.
// Keywords: Aspose.Cells formula delimiter conversion | C# replace semicolon in Excel formula | Aspose.Cells recalculate workbook | SetFormula comma separator | Excel locale delimiter fix .NET | Invalid formula correction Aspose.Cells | Convert semicolon to comma Excel
// Common Searches: Aspose.Cells change formula separator | C# replace ; with , in Excel formula | How to recalculate workbook after fixing formula Aspose | Fix locale‑specific Excel formulas with Aspose.Cells | SetFormula null style parameter meaning
// Developer Intent: Replace semicolons with commas in a formula string, assign it to a cell, trigger calculation, and save the workbook.
// Use Cases: Correct formulas imported from regions that use semicolons before processing with Aspose.Cells. | Normalize user‑entered formulas to the US‑style comma delimiter to avoid calculation errors. | Batch‑process a workbook to replace all semicolon delimiters, recalculate, and export a clean file.
// AI Prompts: Generate C# code that scans every cell in an Aspose.Cells workbook, replaces ';' with ',' in formulas, recalculates, and saves the result. | Explain how to handle formula errors caused by locale‑specific delimiters in Aspose.Cells for .NET, including correction and recalculation steps. | Describe the effect of passing null as the style parameter to SetFormula when updating a corrected formula in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaFixDemo
{
    // Creates a new workbook, writes sample values, converts an invalid formula that uses a semicolon ("=SUM(A1;A2)") to a comma‑separated version, sets the corrected formula in a cell, forces Aspose.Cells to recalculate all formulas, outputs the result, and saves the file as FormulaFixed.xlsx.
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

                // Populate sample values
                cells["A1"].PutValue(5);
                cells["A2"].PutValue(10);

                // Original formula with semicolon (invalid for Aspose.Cells)
                string originalFormula = "=SUM(A1;A2)";
                // Replace semicolons with commas to make it valid
                string correctedFormula = originalFormula.Replace(';', ',');

                // Set the corrected formula; null value lets Aspose.Cells recalculate it
                cells["B1"].SetFormula(correctedFormula, null);

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Output the calculated result
                Console.WriteLine("Calculated value in B1: " + cells["B1"].Value);

                // Save the workbook
                string outputPath = "FormulaFixed.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
