// Title: Copy range formulas and convert to absolute references with Aspose.Cells (.NET C#)
// Description: C# example that creates a workbook, copies formulas from A1:C3 to E1:G3, converts each relative reference to $A$1 style using a regex, recalculates the sheet, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | copy formulas | absolute references | relative to absolute | range copy | Excel formula conversion | SetFormula | Regex | Workbook.CalculateFormula
// Common Searches: Aspose.Cells copy formulas C# | Convert relative references to absolute Aspose.Cells | Copy range with absolute cell addresses .NET | How to set absolute formulas with Aspose.Cells | C# example copy formulas to another range
// Developer Intent: Copy a source range’s formulas to a destination range while turning every relative cell reference into an absolute reference.
// Use Cases: Duplicate a calculation block (e.g., A1:C3) to another area (E1:G3) with fixed cell addresses for consistent results. | Create reusable template sections in financial reports that can be pasted multiple times across a worksheet. | Programmatically migrate formulas between worksheets while preserving absolute references. | Generate multiple report sections that share the same formula logic without manual editing.
// AI Prompts: Write C# code using Aspose.Cells to copy formulas from range A1:C3 to E1:G3 and convert all relative references to $A$1 style. | Provide a reusable method that takes any A1‑style formula string and returns it with absolute references for SetFormula. | Show how to iterate over a source range, copy values and formulas, and convert references with a regular expression in C#. | Explain how to recalculate the workbook after copying absolute formulas with Aspose.Cells.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsFormulaCopyAbsolute
{
    // C# example that creates a workbook, copies formulas from A1:C3 to E1:G3, converts each relative reference to $A$1 style using a regex, recalculates the sheet, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source range with sample data and formulas (relative references)
                // Source range: A1:C3
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Formulas that use relative references
                sheet.Cells["B1"].Formula = "A1*2";   // =A1*2
                sheet.Cells["B2"].Formula = "A2*2";   // =A2*2
                sheet.Cells["B3"].Formula = "A3*2";   // =A3*2
                sheet.Cells["C1"].Formula = "B1+A1"; // =B1+A1
                sheet.Cells["C2"].Formula = "B2+A2"; // =B2+A2
                sheet.Cells["C3"].Formula = "B3+A3"; // =B3+A3

                // Define source and destination ranges
                AsposeRange sourceRange = sheet.Cells.CreateRange("A1:C3");
                AsposeRange destRange = sheet.Cells.CreateRange("E1:G3");

                // Copy values (including formulas) from source to destination
                // We'll handle formula conversion manually
                for (int row = 0; row < sourceRange.RowCount; row++)
                {
                    for (int col = 0; col < sourceRange.ColumnCount; col++)
                    {
                        // Source cell
                        Cell srcCell = sourceRange[row, col];
                        // Destination cell (same offset within destination range)
                        Cell dstCell = destRange[row, col];

                        // Copy value if the cell does not contain a formula
                        if (string.IsNullOrEmpty(srcCell.Formula))
                        {
                            dstCell.PutValue(srcCell.Value);
                        }
                        else
                        {
                            // Convert relative references in the formula to absolute references
                            string absoluteFormula = ConvertToAbsoluteReference(srcCell.Formula);
                            // Set the absolute formula in the destination cell
                            dstCell.SetFormula(absoluteFormula, null);
                        }
                    }
                }

                // Calculate formulas to reflect the new values
                workbook.CalculateFormula();

                // Save the workbook
                workbook.Save("FormulaCopyAbsolute.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Simple conversion: prepend $ to column letters and row numbers for each cell reference
        // This handles basic A1 style references without sheet names or complex ranges.
        private static string ConvertToAbsoluteReference(string formula)
        {
            // Regex to match cell references like A1, B12, AA100 etc.
            // It avoids matching function names or numbers.
            return Regex.Replace(formula, @"(?<![\w$])([A-Z]+)(\d+)", m =>
            {
                string col = m.Groups[1].Value;
                string row = m.Groups[2].Value;
                return $"${col}${row}";
            });
        }
    }
}
