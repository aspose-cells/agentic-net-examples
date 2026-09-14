// Title: Rename Excel worksheets to include their TabId using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, iterates over each Worksheet, and sets its Name to "Sheet_{TabId}". | Show a C# snippet that saves the updated workbook to a new file path, creating the destination folder automatically if it does not exist. | Provide error‑handling logic that checks for a missing input file and logs exceptions while renaming worksheets by TabId.
// Common Searches: aspocells c# rename worksheet to include tabid | how to change Excel sheet names to Sheet_ followed by TabId in .NET | c# Aspose.Cells iterate worksheets and set name based on TabId | save workbook to different folder after renaming sheets with TabId using Aspose.Cells | error handling for missing Excel file when renaming worksheets Aspose.Cells
// Tags: rename worksheets by TabId Aspose.Cells | Aspose.Cells set worksheet name | C# iterate workbook worksheets | save workbook to new path Aspose.Cells | create output directory before saving Excel file

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook with Aspose.Cells, loops through all worksheets, renames each to "Sheet_{TabId}", ensures the output directory exists, and saves the modified file to the specified location with basic error handling.
public static class WorksheetRenamer
{
    /// <param name="inputPath">Full path to the source Excel file.</param>
    /// <param name="outputPath">Full path where the modified Excel file will be saved.</param>
    public static void RenameWorksheetsByTabId(string inputPath, string outputPath)
    {
        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Construct a new name that includes the worksheet's TabId.
                string newName = $"Sheet_{sheet.TabId}";
                sheet.Name = newName;
            }

            // Ensure the output directory exists.
            string outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the modified workbook to the desired location.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Entry point for the console application.
    /// </summary>
    /// <param name="args">Expected arguments: inputPath outputPath</param>
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: WorksheetRenamer <inputPath> <outputPath>");
            return;
        }

        RenameWorksheetsByTabId(args[0], args[1]);
    }
}
