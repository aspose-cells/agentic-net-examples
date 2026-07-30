// Title: C# – Rename Excel worksheets by TabId using Aspose.Cells for .NET
// Description: A console utility that loads an Excel workbook, iterates through all worksheets, builds a name like "Sheet_<TabId>", sanitizes it with CellsHelper.CreateSafeSheetName, assigns the new name, and saves the file. Includes error handling for missing files and runtime exceptions.
// Keywords: Aspose.Cells rename worksheet | TabId sheet name .NET | CreateSafeSheetName example | C# programmatic Excel sheet rename | Excel worksheet safe name | Aspose.Cells command line tool | Workbook worksheet renaming
// Common Searches: how to rename worksheets by TabId Aspose.Cells | C# rename Excel sheet to safe name | Aspose.Cells CreateSafeSheetName usage | rename all sheets in a workbook programmatically | command line tool to change Excel sheet names
// Developer Intent: Automatically give each worksheet a unique, Excel‑compliant name that incorporates its internal TabId.
// Use Cases: Standardize sheet names after importing workbooks from external sources. | Create deterministic identifiers for sheets used in data pipelines or reporting. | Ensure all worksheet names meet Excel length and character restrictions before distribution.
// AI Prompts: Write a C# method that renames every worksheet to "Sheet_<TabId>" using Aspose.Cells and returns a map of old to new names. | Explain why the TabId property is useful for generating unique sheet names in Aspose.Cells. | Generate a PowerShell script that calls the compiled WorksheetRenamer executable with input and output paths.

using System;
using System.IO;
using Aspose.Cells;

// A console utility that loads an Excel workbook, iterates through all worksheets, builds a name like "Sheet_<TabId>", sanitizes it with CellsHelper.CreateSafeSheetName, assigns the new name, and saves the file. Includes error handling for missing files and runtime exceptions.
public class WorksheetRenamer
{
    // Entry point for the console application.
    public static void Main(string[] args)
    {
        // Expect input and output file paths as command‑line arguments.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: WorksheetRenamer <inputPath> <outputPath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            // Verify that the input workbook exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found – {inputPath}");
                return;
            }

            // Perform the renaming operation.
            RenameWorksheetsByTabId(inputPath, outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Renames each worksheet in the workbook to "Sheet_<TabId>"
    // The new name is passed through CellsHelper.CreateSafeSheetName to guarantee it is a valid Excel sheet name.
    public static void RenameWorksheetsByTabId(string inputPath, string outputPath)
    {
        try
        {
            // Load the existing workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Iterate over all worksheets in the workbook.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Create a name based on the internal TabId.
                string proposedName = $"Sheet_{ws.TabId}";

                // Convert to a safe sheet name (handles length limits and illegal characters).
                string safeName = CellsHelper.CreateSafeSheetName(proposedName);

                // Apply the new name.
                ws.Name = safeName;
            }

            // Save the workbook with the updated worksheet names.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Propagate the exception to the caller after logging.
            Console.WriteLine($"Failed to rename worksheets: {ex.Message}");
            throw;
        }
    }
}
