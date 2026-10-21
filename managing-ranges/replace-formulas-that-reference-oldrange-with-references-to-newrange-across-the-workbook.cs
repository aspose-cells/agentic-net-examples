// Title: Replace named range references from 'OldRange' to 'NewRange' in all formulas of an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace named range in all worksheet formulas | how to change OldRange to NewRange in Excel formulas using .NET | bulk update formula text across workbook with Aspose.Cells | C# code to rename a named range inside cell formulas | iterate through cells and modify formula strings Aspose.Cells
// Tags: replace named range references Aspose.Cells | bulk formula text update C# | iterate worksheets modify formulas Aspose.Cells | rename Excel named range in formulas .NET | cell formula string replacement Aspose.Cells | update workbook formulas programmatically C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file with Aspose.Cells, walks through every worksheet and cell, replaces any occurrence of the named range "OldRange" in cell formulas with "NewRange", and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (no password supplied)
            LoadOptions loadOptions = new LoadOptions();
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException ex)
        {
            // Detect password‑protected workbook via message inspection
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Ensure workbook was loaded before proceeding
        if (workbook == null)
        {
            Console.WriteLine("Error: Workbook could not be loaded.");
            return;
        }

        try
        {
            // Iterate through all worksheets and update formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        string updatedFormula = cell.Formula.Replace("OldRange", "NewRange");
                        if (!updatedFormula.Equals(cell.Formula, StringComparison.Ordinal))
                        {
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing worksheets: {ex.Message}");
            return;
        }

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
