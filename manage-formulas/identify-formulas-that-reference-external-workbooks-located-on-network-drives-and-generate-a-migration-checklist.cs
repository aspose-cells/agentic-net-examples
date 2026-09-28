// Title: Identify and list Excel formulas that reference UNC network paths and create a migration checklist with Aspose.Cells for .NET
// AI Prompts: Write a C# program using Aspose.Cells that iterates through every cell in an .xlsx workbook, detects formulas containing UNC network paths, and records the worksheet name, cell address, full formula, and external path. | Create a script that outputs a plain‑text migration checklist summarizing all detected network‑share references, suitable for reviewing before moving files to a new location. | Enhance the solution to parse the sheet and cell reference inside each external link and include those details in the generated checklist.
// Common Searches: C# Aspose.Cells how to list formulas that point to files on a \\ network share | extract external workbook links from an Excel file using Aspose.Cells .NET | generate a migration checklist of Excel cells with UNC path references | scan all worksheets for external references in .xlsx with Aspose.Cells
// Tags: Aspose.Cells scan external UNC references | C# detect network share links in Excel formulas | generate migration checklist for external workbook links | extract external file paths from .xlsx using Aspose.Cells | list cells with external workbook references in .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and cell, checks for formulas, extracts any path enclosed in brackets, and if the path starts with a UNC prefix (\\) records the worksheet name, cell address, full formula, and external path. All findings are written to a text file that serves as a migration checklist for updating or removing network‑share references.
class ExternalFormulaChecker
{
    static void Main(string[] args)
    {
        // Input Excel file path (replace with actual path)
        string inputPath = @"C:\Input\Workbook.xlsx";

        // Output checklist file path
        string outputPath = @"C:\Output\MigrationChecklist.txt";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Ensure the output directory exists
        string outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // List to hold checklist entries
        List<string> checklist = new List<string>();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells in the used range
            foreach (Cell cell in sheet.Cells)
            {
                // Process only cells that contain formulas
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Look for external reference pattern: [path]Sheet!Cell
                    if (formula.Contains("[") && formula.Contains("]"))
                    {
                        int startIdx = formula.IndexOf('[') + 1;
                        int endIdx = formula.IndexOf(']');

                        if (startIdx < endIdx)
                        {
                            string externalReference = formula.Substring(startIdx, endIdx - startIdx);

                            // Identify network drive references (start with \\)
                            if (externalReference.StartsWith(@"\\"))
                            {
                                // Build checklist entry
                                string entry = $"Worksheet: {sheet.Name}, Cell: {cell.Name}, Formula: {formula}, External Path: {externalReference}";
                                checklist.Add(entry);
                            }
                        }
                    }
                }
            }
        }

        try
        {
            // Write the checklist to a text file
            File.WriteAllLines(outputPath, checklist);
            Console.WriteLine($"Checklist written to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write checklist: {ex.Message}");
        }
    }
}
