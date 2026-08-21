// Title: Combine Excel workbooks and preserve unique named ranges with Aspose.Cells for .NET
// Description: Loads a target and a source workbook (creates empty workbooks if files are missing), merges the source into the target using Workbook.Combine, eliminates duplicate named ranges, optionally sorts the remaining names, and saves the merged file.
// Keywords: Aspose.Cells combine workbooks | preserve named ranges | remove duplicate names | sort named ranges | C# Excel merge | Workbook.Combine .NET | handle missing Excel files
// Common Searches: Aspose.Cells merge two workbooks keep named ranges | remove duplicate named ranges after combining Excel files | sort named ranges in merged workbook C# | combine workbooks when source file may not exist Aspose | resolve named range conflicts in Excel using Aspose.Cells
// Developer Intent: Merge two Excel files while keeping each workbook's named ranges distinct and fixing naming conflicts.
// Use Cases: Integrate a template workbook with a data workbook, ensuring both sets of named ranges stay separate. | Consolidate monthly reports into a single file, automatically discarding duplicate range names and ordering the rest for easy navigation. | Run a batch job that merges multiple workbooks, creating placeholder files when some inputs are missing.
// AI Prompts: Generate C# code that uses Aspose.Cells to combine two workbooks, delete duplicate named ranges, and sort the remaining names. | Explain how Workbook.Combine affects named ranges and describe steps to resolve naming conflicts after a merge. | Suggest robust error‑handling patterns for merging workbooks that might be absent on disk.

using System;
using System.IO;
using Aspose.Cells;

// Loads a target and a source workbook (creates empty workbooks if files are missing), merges the source into the target using Workbook.Combine, eliminates duplicate named ranges, optionally sorts the remaining names, and saves the merged file.
class MergeWorkbooksWithNamedRanges
{
    static void Main()
    {
        const string targetPath = "Target.xlsx";
        const string sourcePath = "Source.xlsx";
        const string resultPath = "MergedResult.xlsx";

        Workbook targetWorkbook = null;
        Workbook sourceWorkbook = null;

        try
        {
            // Load target workbook; create a new one if the file does not exist
            if (File.Exists(targetPath))
                targetWorkbook = new Workbook(targetPath);
            else
                targetWorkbook = new Workbook(); // empty workbook

            // Load source workbook; create a new one if the file does not exist
            if (File.Exists(sourcePath))
                sourceWorkbook = new Workbook(sourcePath);
            else
                sourceWorkbook = new Workbook(); // empty workbook

            // Combine source into target
            targetWorkbook.Combine(sourceWorkbook);

            // Remove duplicate named ranges that may have been introduced
            targetWorkbook.Worksheets.Names.RemoveDuplicateNames();

            // Optional: sort the remaining names for a tidy collection
            targetWorkbook.Worksheets.SortNames();

            // Save the merged workbook
            targetWorkbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Merged workbook saved to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during workbook merging:");
            Console.WriteLine(ex.Message);
        }
        finally
        {
            // Dispose workbooks if needed
            targetWorkbook?.Dispose();
            sourceWorkbook?.Dispose();
        }
    }
}
