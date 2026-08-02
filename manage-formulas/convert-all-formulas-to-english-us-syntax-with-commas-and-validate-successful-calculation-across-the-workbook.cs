// Title: Convert Excel formulas to US English syntax (comma separator) and validate calculation with Aspose.Cells for .NET
// Description: Load an .xlsx workbook, set the workbook region to United States so formulas use US English commas, parse any pending formulas, recalculate all cells, detect errors, and save the workbook with the converted formulas.
// Keywords: Aspose.Cells | C# convert formulas | US English formula syntax | comma separator | region USA | parse formulas | calculate workbook | formula error detection | Excel formula conversion .NET | set workbook locale
// Common Searches: Aspose.Cells change formula separator to comma | set workbook region to United States Aspose.Cells | validate formula calculation after locale change .NET | convert Excel formulas to US English using Aspose.Cells | parse and recalculate formulas in C# Aspose.Cells
// Developer Intent: Ensure all formulas in an Excel file are converted to US English syntax with commas and verify they compute correctly using Aspose.Cells.
// Use Cases: Batch‑process international Excel files to standardize formulas for US‑based reporting. | Pre‑process workbooks before importing into a US‑centric analytics platform. | Run an automated quality check in CI pipelines to catch formula errors after locale conversion. | Integrate into a data migration tool that normalizes formula syntax across regions.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, sets Region = CountryCode.USA, parses formulas, calculates them, and throws an exception for any error cells. | Show how to iterate through all worksheets and cells to confirm no formula errors after converting to US English syntax. | Explain the impact of workbook.Settings.Region on formula parsing and calculation in Aspose.Cells. | Provide a PowerShell script that invokes a .NET assembly to perform formula conversion and validation for Excel files.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaConversion
{
    // Load an .xlsx workbook, set the workbook region to United States so formulas use US English commas, parse any pending formulas, recalculate all cells, detect errors, and save the workbook with the converted formulas.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with actual path)
            string inputPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure formulas are interpreted using English (US) locale (comma as argument separator)
            workbook.Settings.Region = CountryCode.USA;

            // Parse any formulas that were set without immediate parsing
            // false -> do not ignore errors while parsing
            workbook.ParseFormulas(false);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Validate that every formula cell has been calculated without error
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        // If the cell result is an error, throw an exception
                        if (cell.Type == CellValueType.IsError)
                        {
                            throw new InvalidOperationException(
                                $"Formula error in sheet \"{sheet.Name}\", cell {cell.Name}");
                        }
                    }
                }
            }

            // Save the workbook with converted formulas
            string outputPath = "output_converted.xlsx";
            workbook.Save(outputPath);
        }
    }
}
