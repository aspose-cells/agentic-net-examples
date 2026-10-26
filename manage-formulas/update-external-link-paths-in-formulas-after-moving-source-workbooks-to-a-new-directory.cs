// Title: Update external workbook link paths in Excel formulas after moving source files using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace old external link folder in formula strings | programmatically change external reference paths in Excel workbook with Aspose.Cells | update formula external links after moving source directory using .NET | bulk modify Excel formulas to point to new source folder in C# | how to adjust external workbook paths in formulas with Aspose.Cells API
// Tags: update external link paths in Excel formulas Aspose.Cells | replace folder path in formula references C# | bulk modify formula external references .NET | adjust workbook external links after folder move | iterate cells to edit formula strings Aspose.Cells | external link directory update Excel workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkUpdater
{
    // The example loads a target workbook, defines the previous and new directories for external sources, iterates through every worksheet and cell, replaces occurrences of the old directory in each formula string with the new directory, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths for the target workbook and the updated output
                string sourcePath = @"C:\Workbooks\TargetWorkbook.xlsx";
                string outputPath = @"C:\Workbooks\TargetWorkbook_Updated.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source workbook not found: {sourcePath}");
                    return;
                }

                // Load the workbook that contains formulas with external links
                Workbook workbook = new Workbook(sourcePath);

                // Define the old and new directories for external links
                string oldDirectory = @"C:\OldSourceFolder\";
                string newDirectory = @"D:\NewSourceFolder\";

                // Update external link paths in all formulas
                UpdateExternalLinkPaths(workbook, oldDirectory, newDirectory);

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void UpdateExternalLinkPaths(Workbook workbook, string oldDir, string newDir)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;
                        if (!string.IsNullOrEmpty(formula) && formula.Contains(oldDir, StringComparison.OrdinalIgnoreCase))
                        {
                            string updatedFormula = formula.Replace(oldDir, newDir, StringComparison.OrdinalIgnoreCase);
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }
        }
    }
}
