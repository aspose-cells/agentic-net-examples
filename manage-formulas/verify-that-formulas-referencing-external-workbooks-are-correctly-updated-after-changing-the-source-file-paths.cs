// Title: Update external workbook paths in Excel formulas and verify the changes with Aspose.Cells for .NET
// AI Prompts: Replace all occurrences of an old external workbook path with a new path in every formula cell of an Excel file using Aspose.Cells for C#. | Scan the workbook after the replacement and list any cells whose formulas still contain the original path. | Save the modified workbook to a new file while handling load and save exceptions gracefully. | Iterate through all worksheets and cells to detect and update external link references in formulas.
// Common Searches: Aspose.Cells replace external link path in Excel formula C# | verify external workbook references updated after path change Aspose.Cells | C# code to update external workbook references in Excel formulas | how to check for leftover old paths in Excel formulas using Aspose.Cells | update external links in formulas and save workbook with Aspose.Cells .NET
// Tags: update external links in formulas Aspose.Cells | replace workbook path in Excel formulas C# | validate external reference updates Aspose.Cells | iterate cells to modify formula strings C# | save workbook with updated external references Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, replaces a specified old external workbook path with a new one in every formula cell, verifies that no formula still contains the old path (reporting any that do), and saves the updated workbook to a new file, handling load and save errors using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputFile = "WorkbookWithExternalLinks.xlsx";
        const string outputFile = "WorkbookWithExternalLinks_Updated.xlsx";

        // Verify input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        Workbook wb;
        try
        {
            wb = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Define old and new external workbook paths
        string oldPath = @"C:\OldPath\Source.xlsx";
        string newPath = @"D:\NewPath\Source.xlsx";

        // Update external references manually by replacing the old path in formulas
        foreach (Worksheet sheet in wb.Worksheets)
        {
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                {
                    if (cell.Formula.Contains(oldPath))
                    {
                        string updatedFormula = cell.Formula.Replace(oldPath, newPath);
                        cell.Formula = updatedFormula;
                    }
                }
            }
        }

        // Verify that all formulas have been updated
        bool allUpdated = true;
        foreach (Worksheet sheet in wb.Worksheets)
        {
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                {
                    if (cell.Formula.Contains(oldPath))
                    {
                        allUpdated = false;
                        Console.WriteLine($"Cell {sheet.Name}!{cell.Name} still contains the old path.");
                    }
                }
            }
        }

        Console.WriteLine(allUpdated
            ? "All external references were successfully updated."
            : "Some external references were not updated.");

        // Save the updated workbook
        try
        {
            wb.Save(outputFile);
            Console.WriteLine($"Workbook saved as {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
