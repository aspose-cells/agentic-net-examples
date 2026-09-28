// Title: Convert semicolon‑separated Excel formula arguments to commas and recalculate workbook with Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook using Aspose.Cells, replace every ';' in cell formulas with ',', then invoke Workbook.CalculateFormula to update results. | Assign Workbook.Settings.CultureInfo = new CultureInfo("en-US") so that Aspose.Cells parses formulas with comma separators before saving. | Traverse all worksheets and cells, use Cell.IsFormula to identify formulas, modify the Formula string, and finally call Workbook.CalculateFormula.
// Common Searches: aspnet replace semicolon in Excel formula arguments using Aspose.Cells | how to change formula delimiter from ';' to ',' in Aspose.Cells workbook | recalculate formulas after editing them with Aspose.Cells .NET | set workbook culture to en-US for formula parsing Aspose.Cells | iterate over cells to update formulas in Aspose.Cells example
// Tags: formula argument separator conversion Aspose.Cells | workbook cultureinfo en-us setting Aspose.Cells | bulk formula update across worksheets Aspose.Cells | recalculate workbook formulas Aspose.Cells | locale-aware formula parsing Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, forces the workbook to use the en-US locale so commas are the argument separator, scans every worksheet and cell, replaces any semicolons in formulas with commas, recalculates all formulas, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Force the workbook to use a culture that expects commas as argument separators
            workbook.Settings.CultureInfo = new CultureInfo("en-US");

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Use IsFormula to check if the cell contains a formula
                        if (cell.IsFormula)
                        {
                            // Replace any semicolons with commas in the formula
                            string formula = cell.Formula;
                            if (formula.Contains(";"))
                            {
                                cell.Formula = formula.Replace(";", ",");
                            }
                        }
                    }
                }
            }

            // Recalculate all formulas to ensure they evaluate correctly
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
