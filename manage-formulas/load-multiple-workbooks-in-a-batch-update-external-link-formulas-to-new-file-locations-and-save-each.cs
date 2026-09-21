// Title: Batch update external link formulas in multiple Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that loads every .xlsx file in a folder, replaces a specified old base path in external link formulas with a new path, and saves the modified workbooks to another directory. | Add comprehensive error handling to the batch updater so that it logs files that fail to load or save and returns a summary of total, successful, and failed workbook counts. | Refactor the solution into a reusable method that accepts input folder, output folder, old base path, and new base path parameters and returns the total number of formulas updated across all workbooks.
// Common Searches: how to change external link paths in formulas for many Excel files using Aspose.Cells C# | batch replace folder path in external references across multiple workbooks .NET | Aspose.Cells iterate all cells to modify formulas in a directory of .xlsx files | C# program to update external workbook links after moving files to a new location
// Tags: Aspose.Cells batch external link path update | C# replace formula base path in Excel workbooks | process multiple .xlsx files with Aspose.Cells | update external references programmatically .NET | save modified workbooks to separate folder

using System;
using System.IO;
using Aspose.Cells;

// The example loads each .xlsx file from a source directory, walks through every worksheet and cell, replaces occurrences of a given old base path in external‑link formulas with a new base path, and saves the updated workbooks to an output directory using Aspose.Cells for .NET.
class ExternalLinkUpdater
{
    // Update external link formulas in all workbooks within a folder.
    static void Main()
    {
        // Folder containing the source workbooks.
        string inputFolder = @"C:\SourceWorkbooks";

        // Folder where the updated workbooks will be saved.
        string outputFolder = @"C:\UpdatedWorkbooks";

        // Old and new base paths used in external link formulas.
        string oldBasePath = @"C:\OldFolder\";
        string newBasePath = @"D:\NewFolder\";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Verify the input folder exists.
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        string[] workbookFiles;
        try
        {
            // Get all Excel files in the input folder.
            workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing input folder: {ex.Message}");
            return;
        }

        foreach (string filePath in workbookFiles)
        {
            // Verify the file exists before attempting to load.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each cell that contains a formula.
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Replace the old external link path with the new one.
                            if (!string.IsNullOrEmpty(formula) && formula.Contains(oldBasePath))
                            {
                                string updatedFormula = formula.Replace(oldBasePath, newBasePath);
                                cell.Formula = updatedFormula;
                            }
                        }
                    }
                }

                // Determine the output file path.
                string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the modified workbook.
                workbook.Save(outputFilePath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("External link update completed for all workbooks.");
    }
}
