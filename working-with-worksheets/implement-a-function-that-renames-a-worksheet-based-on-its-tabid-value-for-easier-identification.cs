// Title: Rename Excel worksheets by TabId with Aspose.Cells for .NET (C#)
// Common Searches: C# code to rename Excel sheets using their TabId with Aspose.Cells | How to change worksheet names to include TabId in Aspose.Cells .NET | Programmatically set worksheet name based on TabId in a workbook using Aspose.Cells | Aspose.Cells example for iterating worksheets and updating names by TabId | Rename Excel worksheet names to Sheet_{TabId} using Aspose.Cells C#
// Tags: rename worksheet by TabId Aspose.Cells | set worksheet name programmatically C# | Aspose.Cells workbook worksheet iteration | Excel sheet name length limit handling Aspose | load and save workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetTabIdRenamer
{
    // The sample loads an Excel workbook with Aspose.Cells, iterates through each worksheet, and renames it to "Sheet_{TabId}" (truncated to Excel's 31‑character limit). It logs the original name and TabId, then saves the modified workbook to a new file.
    class Program
    {
        /// <param name="workbook">The workbook whose worksheets will be renamed.</param>
        static void RenameWorksheetsByTabId(Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Preserve the original name for logging
                string originalName = sheet.Name;
                int tabId = sheet.TabId;

                // Construct a new name that includes the TabId.
                // Ensure the name does not exceed Excel's 31‑character limit.
                string newName = $"Sheet_{tabId}";
                if (newName.Length > 31)
                {
                    newName = newName.Substring(0, 31);
                }

                // Apply the new name.
                sheet.Name = newName;

                Console.WriteLine($"Renamed worksheet \"{originalName}\" (TabId={tabId}) to \"{newName}\".");
            }
        }

        static void Main(string[] args)
        {
            // Expect at least one argument: the path to the input workbook.
            // Optional second argument: the path for the output workbook.
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: WorksheetTabIdRenamer <inputWorkbookPath> [outputWorkbookPath]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args.Length >= 2 ? args[1] : Path.Combine(
                Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory(),
                Path.GetFileNameWithoutExtension(inputPath) + "_Renamed" + Path.GetExtension(inputPath));

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{inputPath}\" with {workbook.Worksheets.Count} worksheet(s).");

                // Rename worksheets based on their TabId.
                RenameWorksheetsByTabId(workbook);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
