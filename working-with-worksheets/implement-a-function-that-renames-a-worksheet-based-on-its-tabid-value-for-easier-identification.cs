// Title: Rename Excel worksheets by TabId using Aspose.Cells for .NET
// Description: C# code that loads a workbook, iterates through each worksheet, builds a name containing the sheet's TabId (e.g., "Sheet_5"), sanitizes it with CellsHelper.CreateSafeSheetName to meet Excel naming rules, assigns the new name, and saves the file.
// Keywords: Aspose.Cells rename worksheet TabId | C# rename Excel sheets by TabId | CreateSafeSheetName Aspose.Cells | worksheet naming rules Excel .NET | programmatic sheet renaming Aspose
// Common Searches: how to rename worksheets by TabId Aspose.Cells | Aspose.Cells C# change sheet name to include TabId | safe sheet name generation Aspose.Cells | rename all Excel sheets programmatically .NET | Excel TabId based sheet naming example
// Developer Intent: Update every worksheet name to include its TabId while ensuring the name complies with Excel constraints.
// Use Cases: Standardize sheet identifiers for automated processing pipelines. | Create traceable workbook versions where sheet order is reflected in the name. | Guarantee compliance with Excel's 31‑character limit and prohibited characters after bulk renaming.
// AI Prompts: Write C# code with Aspose.Cells that renames each worksheet to "Sheet_{TabId}" and logs the original and new names. | Show how to handle duplicate names when using CellsHelper.CreateSafeSheetName while renaming sheets by TabId. | Explain the sanitization steps performed by CellsHelper.CreateSafeSheetName and which characters are removed.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# code that loads a workbook, iterates through each worksheet, builds a name containing the sheet's TabId (e.g., "Sheet_5"), sanitizes it with CellsHelper.CreateSafeSheetName to meet Excel naming rules, assigns the new name, and saves the file.
    public static class WorksheetRenamer
    {
        /// <param name="filePath">Full path of the workbook to process.</param>
        public static void RenameWorksheetsByTabId(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the existing workbook from the specified file.
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets in the collection.
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Build a name that contains the TabId.
                    string proposedName = $"Sheet_{ws.TabId}";

                    // Ensure the name is valid for Excel (max 31 chars, no illegal characters).
                    string safeName = CellsHelper.CreateSafeSheetName(proposedName);

                    // Assign the safe name to the worksheet.
                    ws.Name = safeName;
                }

                // Save the workbook, overwriting the original file.
                workbook.Save(filePath);
                Console.WriteLine("Worksheet names updated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the console application.
    public class Program
    {
        public static void Main(string[] args)
        {
            // Expect a file path argument; otherwise use a default placeholder.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            WorksheetRenamer.RenameWorksheetsByTabId(filePath);
        }
    }
}
