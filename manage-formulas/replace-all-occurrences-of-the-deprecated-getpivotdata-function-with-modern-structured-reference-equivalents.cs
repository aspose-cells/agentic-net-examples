// Title: Replace GETPIVOTDATA formulas with structured references in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace GETPIVOTDATA function with structured reference | convert GETPIVOTDATA formulas to Sheet[Field] syntax using Aspose.Cells | programmatically update Excel pivot formulas to structured references in .NET | regex to find and replace GETPIVOTDATA in a workbook with Aspose.Cells | load workbook, iterate cells, modify formulas Aspose.Cells C# example
// Tags: convert GETPIVOTDATA to structured reference Aspose.Cells | modify Excel formulas programmatically C# | regex based formula transformation Aspose.Cells | iterate worksheets and cells formulas .NET | load and save Excel workbook Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, scans every worksheet and cell for formulas containing the deprecated GETPIVOTDATA function, uses a regular expression to capture the field name and pivot sheet, replaces each occurrence with a structured reference like =SheetName[FieldName], updates the cell formula when changed, and saves the modified workbook to a new file.
class ReplaceGetPivotData
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook (add password here if required)
            LoadOptions loadOptions = new LoadOptions();
            // loadOptions.Password = "your_password"; // uncomment and set if needed
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

        // Regex to capture GETPIVOTDATA arguments:
        // Example: GETPIVOTDATA("Sales",PivotTable!$A$3)
        // Captures field name (group 1) and pivot table sheet name (group 2)
        Regex getPivotRegex = new Regex(
            @"GETPIVOTDATA\(\s*""([^""]+)""\s*,\s*([A-Za-z0-9_]+)![$]?[A-Za-z]+\d+\s*\)",
            RegexOptions.IgnoreCase);

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Check if the formula contains GETPIVOTDATA
                        if (formula.IndexOf("GETPIVOTDATA", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace each GETPIVOTDATA occurrence with a structured reference
                            string newFormula = getPivotRegex.Replace(formula, match =>
                            {
                                // Extract field name and pivot table sheet name
                                string fieldName = match.Groups[1].Value;
                                string pivotSheetName = match.Groups[2].Value;

                                // Build a simple structured reference: =PivotSheetName[FieldName]
                                return $"={pivotSheetName}[{fieldName}]";
                            });

                            // Update the cell formula only if a change occurred
                            if (!formula.Equals(newFormula, StringComparison.Ordinal))
                            {
                                cell.Formula = newFormula;
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
