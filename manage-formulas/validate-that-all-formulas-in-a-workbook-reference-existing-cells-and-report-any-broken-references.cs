// Title: C# – Detect Broken Formula References in Excel with Aspose.Cells
// Description: Loads an Excel workbook, parses all formulas, examines each precedent, and reports references to missing worksheets, invalid ranges, or external files. Outputs detailed messages to the console for quick cleanup.
// Keywords: Aspose.Cells | C# | .NET | formula validation | broken references | Excel precedents | missing worksheet | invalid range | external link detection | cell reference check
// Common Searches: Aspose.Cells find broken formulas | C# check Excel formula references | detect missing worksheet in formula Aspose | validate Excel formulas .NET | list external links in workbook Aspose.Cells
// Developer Intent: Identify and list any formula references that point to non‑existent sheets, out‑of‑range cells, or external workbooks.
// Use Cases: Clean up workbooks before publishing or sharing. | Automate formula integrity checks in CI/CD pipelines. | Generate audit reports of broken references after data migration. | Alert users to external links that require updating or embedding. | Prevent runtime errors caused by corrupted formula ranges.
// AI Prompts: Write C# code using Aspose.Cells that returns a collection of broken formula reference messages, handling missing sheets, invalid ranges, and external links. | Enhance the validation loop to also flag references that exceed the worksheet's MaxDataRow or MaxDataColumn limits. | Create a reusable method `IEnumerable<string> GetBrokenFormulaReferences(Workbook wb)` that encapsulates all checks and can be called from unit tests. | Suggest how to integrate this validation into a GitHub Actions workflow that fails the build on any broken reference. | Provide a version of the script that logs results to a JSON file for downstream processing.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaValidator
{
    // Loads an Excel workbook, parses all formulas, examines each precedent, and reports references to missing worksheets, invalid ranges, or external files. Outputs detailed messages to the console for quick cleanup.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Parse all formulas (ignore errors to continue processing)
                workbook.ParseFormulas(false);

                // List to hold broken reference descriptions
                List<string> brokenReferences = new List<string>();

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Determine the range to scan (including cells that may contain formulas)
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    // Scan all cells within the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Process only formula cells
                            if (cell.IsFormula)
                            {
                                // Get all precedents (references) used in the formula
                                ReferredAreaCollection precedents = cell.GetPrecedents();

                                // If there are no precedents, continue
                                if (precedents == null) continue;

                                // Examine each referenced area
                                foreach (ReferredArea area in precedents)
                                {
                                    // Handle external links – detailed validation of external files is omitted
                                    if (area.IsExternalLink)
                                    {
                                        brokenReferences.Add(
                                            $"Cell {cell.Name} in sheet \"{sheet.Name}\" references external link \"{area.ExternalFileName}\" which cannot be validated.");
                                        continue;
                                    }

                                    // Determine the target worksheet
                                    Worksheet targetSheet = sheet; // default to current sheet
                                    if (!string.IsNullOrEmpty(area.SheetName))
                                    {
                                        targetSheet = workbook.Worksheets[area.SheetName];
                                        if (targetSheet == null)
                                        {
                                            brokenReferences.Add(
                                                $"Cell {cell.Name} in sheet \"{sheet.Name}\" references non‑existent sheet \"{area.SheetName}\".");
                                            continue; // cannot validate further without a sheet
                                        }
                                    }

                                    // Validate row and column indices (they must be non‑negative)
                                    if (area.StartRow < 0 || area.StartColumn < 0 ||
                                        area.EndRow < 0 || area.EndColumn < 0)
                                    {
                                        brokenReferences.Add(
                                            $"Cell {cell.Name} in sheet \"{sheet.Name}\" has an invalid reference range.");
                                        continue;
                                    }

                                    // Additional range checks could be added here if needed
                                }
                            }
                        }
                    }
                }

                // Report results
                if (brokenReferences.Count == 0)
                {
                    Console.WriteLine("No broken formula references were found.");
                }
                else
                {
                    Console.WriteLine("Broken formula references detected:");
                    foreach (string msg in brokenReferences)
                    {
                        Console.WriteLine("- " + msg);
                    }
                }

                // Save the workbook if any modifications were made (optional)
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
