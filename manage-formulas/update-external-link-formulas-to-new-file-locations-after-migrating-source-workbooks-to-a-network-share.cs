// Title: Replace old UNC base path with new network share in external link formulas of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook with Aspose.Cells, scan all cells for formulas that contain a specific UNC folder path, replace the old path with a new UNC path, and save the workbook. | Iterate through each worksheet and cell, detect external link formulas referencing a legacy network share, update the reference to a new share, and preserve existing formula logic. | Add robust error handling to verify the source workbook exists before performing bulk UNC path replacement in external formulas with Aspose.Cells.
// Common Searches: C# Aspose.Cells replace old UNC path in external link formulas | how to update external workbook references after moving files to a new network share using Aspose.Cells | programmatically change Excel external link paths in .NET | bulk edit formulas containing network share paths with Aspose.Cells
// Tags: replace UNC base path in Aspose.Cells formulas | bulk update external link references .NET | migrate Excel external links to new network share | Aspose.Cells error handling for missing workbook | iterate cells to modify formulas Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads Summary.xlsx, iterates through every worksheet and cell, replaces any formula that contains the old UNC base path with the new UNC base path, saves the modified workbook as Summary_Updated.xlsx, and includes error handling for a missing source file.
class UpdateExternalLinks
{
    static void Main()
    {
        // Paths
        string workbookPath = @"C:\Reports\Summary.xlsx";
        string outputPath = @"C:\Reports\Summary_Updated.xlsx";

        // Old and new base paths for the source workbooks
        string oldBasePath = @"\\oldserver\share\SourceFiles\";
        string newBasePath = @"\\newserver\share\SourceFiles\";

        try
        {
            // Verify the input workbook exists
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook not found: {workbookPath}");

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Update formulas that embed the old path directly
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) &&
                        cell.Formula.IndexOf(oldBasePath, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the old path with the new path inside the formula
                        cell.Formula = cell.Formula.Replace(oldBasePath, newBasePath, StringComparison.OrdinalIgnoreCase);
                    }
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook updated and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
