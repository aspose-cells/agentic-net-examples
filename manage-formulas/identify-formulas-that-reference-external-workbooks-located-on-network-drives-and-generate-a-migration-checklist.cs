// Title: Detect external workbook links in Excel formulas and create a migration checklist with Aspose.Cells for .NET
// Common Searches: C# code to list external workbook references in an Excel file using Aspose.Cells | how to generate a checklist for migrating network‑drive linked Excel formulas in .NET | Aspose.Cells find formulas that reference other workbooks on a shared drive | extract and replace external workbook paths in Excel formulas with C# | scan Excel workbook for bracketed external references and output report Aspose.Cells | automate migration of Excel external links to local files using Aspose.Cells
// Tags: Aspose.Cells external reference extraction | Excel formula link enumeration with C# | migration checklist generation for external workbook links | detect network drive paths in Excel formulas | C# workbook scanning for external references

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ExternalFormulaMigration
{
    // The sample loads an .xlsx workbook with Aspose.Cells, iterates through each worksheet's used range, identifies cells whose formulas contain bracketed workbook references (typically network‑drive links), collects unique external references, and builds a text‑based migration checklist that outlines verification, copying, formula updating, and testing steps. The checklist is printed to the console and saved as a text file alongside the source workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the workbook to be analyzed
                string workbookPath = @"C:\Path\To\YourWorkbook.xlsx";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook not found: {workbookPath}");
                    return;
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(workbookPath);

                // Store unique external references found in formulas
                HashSet<string> externalReferences = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet (explicitly use Aspose.Cells.Range to avoid ambiguity)
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                    if (usedRange == null)
                        continue;

                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Iterate through each cell in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Process only cells that contain a formula
                            if (cell.IsFormula)
                            {
                                string formula = cell.Formula;

                                // Check for external workbook reference (enclosed in square brackets)
                                if (formula.Contains("[") && formula.Contains("]"))
                                {
                                    int startIdx = formula.IndexOf('[');
                                    int endIdx = formula.IndexOf(']', startIdx);
                                    if (startIdx >= 0 && endIdx > startIdx)
                                    {
                                        string reference = formula.Substring(startIdx + 1, endIdx - startIdx - 1);
                                        externalReferences.Add(reference);
                                    }
                                }
                            }
                        }
                    }
                }

                // Build migration checklist
                List<string> checklist = new List<string>
                {
                    "=== External Workbook Migration Checklist ===",
                    $"Date Generated: {DateTime.Now}",
                    $"Source Workbook: {Path.GetFileName(workbookPath)}",
                    ""
                };

                if (externalReferences.Count == 0)
                {
                    checklist.Add("No external workbook references were found.");
                }
                else
                {
                    int idx = 1;
                    foreach (string extRef in externalReferences)
                    {
                        checklist.Add($"{idx}. External Reference: {extRef}");
                        checklist.Add("   - Verify accessibility of the network location.");
                        checklist.Add("   - Copy the referenced workbook to the target environment (if required).");
                        checklist.Add("   - Update the formula to point to the new location or embed the data.");
                        checklist.Add("   - Test the formula after migration to ensure correct results.");
                        checklist.Add("");
                        idx++;
                    }
                }

                // Output the checklist to console
                foreach (string line in checklist)
                {
                    Console.WriteLine(line);
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(workbookPath) ?? "";
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Write the checklist to a text file
                string outputPath = Path.Combine(outputDir, "MigrationChecklist.txt");
                File.WriteAllLines(outputPath, checklist);
                Console.WriteLine($"Checklist saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
