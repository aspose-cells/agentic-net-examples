using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

namespace AsposeCellsExamples
{
    public class ReplaceGetPivotData
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook containing GETPIVOTDATA formulas
                Workbook workbook = new Workbook(inputPath);

                // Regex to capture GETPIVOTDATA arguments
                Regex getPivotRegex = new Regex(
                    @"GETPIVOTDATA\s*\(\s*([^,]+)\s*,\s*([^,\)]+)\s*(.*)\)",
                    RegexOptions.IgnoreCase);

                // Iterate through worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through cells with formulas
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (!cell.IsFormula) continue;

                        string formula = cell.Formula;
                        Match match = getPivotRegex.Match(formula);
                        if (!match.Success) continue; // Not a GETPIVOTDATA formula

                        // Extract arguments
                        string dataFieldArg = match.Groups[1].Value.Trim();   // e.g. "Sum of Sales"
                        string pivotRefArg  = match.Groups[2].Value.Trim();   // e.g. $A$3

                        // Remove surrounding quotes from data field if present
                        string dataField = dataFieldArg.Trim('\"');

                        // Clean cell reference (remove $ signs) and obtain the cell
                        string cleanRef = pivotRefArg.Replace("$", string.Empty);
                        Cell pivotRefCell = sheet.Cells[cleanRef];
                        if (pivotRefCell == null) continue; // Unable to resolve reference

                        // Find the pivot table that contains the reference cell
                        string pivotTableName = null;
                        foreach (PivotTable pt in sheet.PivotTables)
                        {
                            CellArea range = pt.TableRange1;
                            if (pivotRefCell.Row >= range.StartRow && pivotRefCell.Row <= range.EndRow &&
                                pivotRefCell.Column >= range.StartColumn && pivotRefCell.Column <= range.EndColumn)
                            {
                                pivotTableName = pt.Name;
                                break;
                            }
                        }

                        if (string.IsNullOrEmpty(pivotTableName)) continue; // Pivot table not found

                        // Build structured reference: =PivotTableName[DataField]
                        string newFormula = $"={pivotTableName}[{dataField}]";

                        // Replace old formula
                        cell.Formula = newFormula;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceGetPivotData.Run();
        }
    }
}