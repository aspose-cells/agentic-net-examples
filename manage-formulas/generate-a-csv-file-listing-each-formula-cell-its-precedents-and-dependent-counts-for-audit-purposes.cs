// Title: Create a CSV audit of Excel formula cells with precedent and dependent counts using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that loads an .xlsx workbook with Aspose.Cells, iterates every worksheet, extracts each formula cell, counts its precedents via GetPrecedents and its dependents via GetDependents(true), and writes SheetName, CellAddress, Formula, PrecedentCount, and DependentCount to a CSV file. | Extend the program to also output the addresses of all precedent cells in the CSV after the precedent count column. | Add comprehensive error handling around GetPrecedents and GetDependents so that any exception is logged to the console and the audit continues processing remaining cells.
// Common Searches: aspnet generate CSV of Excel formula precedents and dependents using Aspose.Cells | c# code to list formula cells with dependency counts in an .xlsx workbook | how to export a formula audit report from Aspose.Cells to CSV | retrieve cell dependents recursively with Aspose.Cells C# example
// Tags: aspocells export formula audit to csv | c# get formula precedents aspocells | c# retrieve dependent cells aspocells | excel formula dependency report aspocells | csv generation from workbook formulas .net

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, scans each worksheet for formula cells, uses GetPrecedents to count referenced cells and GetDependents(true) to count cells that reference the formula, and writes a CSV containing SheetName, CellAddress, Formula, PrecedentCount, and DependentCount for audit purposes.
class FormulaAudit
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path for the generated CSV report
        string outputCsv = "formula_audit.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        try
        {
            // Load the workbook (Aspose.Cells handles all formats)
            Workbook workbook = new Workbook(inputPath);

            // Ensure the directory for the CSV exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputCsv));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Prepare CSV writer
            using (StreamWriter writer = new StreamWriter(outputCsv))
            {
                // Write CSV header
                writer.WriteLine("SheetName,CellAddress,Formula,PrecedentCount,DependentCount");

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all cells in the sheet
                    foreach (Cell cell in cells)
                    {
                        // Process only formula cells
                        if (cell.IsFormula)
                        {
                            int precedentCount = 0;
                            int dependentCount = 0;

                            // Retrieve precedents (cells referenced by this formula)
                            try
                            {
                                ReferredAreaCollection precedents = cell.GetPrecedents();
                                precedentCount = precedents?.Count ?? 0;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to get precedents for {cell.Name}: {ex.Message}");
                            }

                            // Retrieve dependents (cells that reference this cell)
                            // 'true' gets all dependents recursively
                            try
                            {
                                Cell[] dependents = cell.GetDependents(true);
                                dependentCount = dependents?.Length ?? 0;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to get dependents for {cell.Name}: {ex.Message}");
                            }

                            // Write a line to the CSV file
                            string line = string.Format("{0},{1},{2},{3},{4}",
                                sheet.Name,
                                cell.Name,
                                cell.Formula,
                                precedentCount,
                                dependentCount);

                            writer.WriteLine(line);
                        }
                    }
                }
            }

            // Inform the user of successful generation
            Console.WriteLine("Formula audit CSV generated at: " + Path.GetFullPath(outputCsv));
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine("An error occurred during processing:");
            Console.WriteLine(ex.Message);
        }
    }
}
