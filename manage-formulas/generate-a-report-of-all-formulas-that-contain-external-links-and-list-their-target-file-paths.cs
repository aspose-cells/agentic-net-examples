// Title: Report Excel formulas with external links and their file paths – Aspose.Cells C#
// Description: Loads an Excel workbook, iterates every worksheet and used cell, identifies formulas that reference external workbooks, gathers each linked file path via GetPrecedents or the workbook's ExternalLinks collection, and prints a concise report. Shows how to use Aspose.Cells for .NET to audit cross‑file references.
// Keywords: Aspose.Cells external link detection | C# list Excel formula references | GetPrecedents external workbook | Excel cross‑file audit | report external file paths
// Common Searches: Aspose.Cells find formulas that reference other workbooks | C# extract external workbook paths from Excel formulas | list cells with external links using Aspose.Cells | generate external link report .NET
// Developer Intent: Create a concise report of all formulas that reference external workbooks and show each target file path.
// Use Cases: Audit workbook dependencies before distribution | Document data sources for compliance or migration | Identify broken or outdated external references
// AI Prompts: Generate a method that returns a dictionary mapping cell addresses to arrays of linked file paths using Aspose.Cells. | Adapt the sample to export the external‑link report to CSV or JSON. | Add error handling to skip cells without precedents and log missing external‑link collections.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, iterates every worksheet and used cell, identifies formulas that reference external workbooks, gathers each linked file path via GetPrecedents or the workbook's ExternalLinks collection, and prints a concise report. Shows how to use Aspose.Cells for .NET to audit cross‑file references.
class ExternalLinkReport
{
    static void Main()
    {
        // Load the workbook (replace with actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // List to hold report entries
        List<string> reportLines = new List<string>();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Check if the cell is a formula that contains an external link
                if (cell.ContainsExternalLink)
                {
                    // Collect external file names referenced by this formula
                    HashSet<string> externalFiles = new HashSet<string>();

                    // Get all precedents (references) of the formula
                    ReferredAreaCollection precedents = cell.GetPrecedents();
                    if (precedents != null)
                    {
                        foreach (ReferredArea area in precedents)
                        {
                            if (area.IsExternalLink)
                            {
                                externalFiles.Add(area.ExternalFileName);
                            }
                        }
                    }

                    // If no external files were found via precedents, fall back to the workbook's external links collection
                    if (externalFiles.Count == 0 && workbook.Worksheets.ExternalLinks.Count > 0)
                    {
                        foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
                        {
                            externalFiles.Add(link.DataSource);
                        }
                    }

                    // Build the report line
                    string files = string.Join(", ", externalFiles);
                    string line = $"{sheet.Name}!{cell.Name}: {cell.Formula} -> {files}";
                    reportLines.Add(line);
                }
            }
        }

        // Output the report
        Console.WriteLine("Formulas containing external links:");
        foreach (string line in reportLines)
        {
            Console.WriteLine(line);
        }

        // Save the workbook (no modifications made, just demonstrating the save rule)
        workbook.Save("output.xlsx");
    }
}
