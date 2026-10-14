// Title: Rename Excel worksheets to unique names based on their TabId using Aspose.Cells for .NET
// AI Prompts: Generate a C# routine that opens an Excel workbook with Aspose.Cells, reads each worksheet's TabId, creates a name like "Sheet_<TabId>", ensures the name does not clash with existing sheets, renames the worksheet, and saves the file. | Write code that iterates through Workbook.Worksheets, retrieves Worksheet.TabId, builds a unique sheet name, assigns it to Worksheet.Name, and includes error handling for missing files and duplicate names. | Create a method that accepts input and output file paths, loads the workbook, renames all sheets based on their TabId while handling name collisions, and returns a status message.
// Common Searches: Aspose.Cells C# rename all worksheets using TabId and keep names unique | How to programmatically set Excel sheet name to its TabId with Aspose.Cells .NET | C# example for iterating worksheets and renaming them based on TabId in Aspose.Cells
// Tags: Aspose.Cells TabId sheet naming pattern | avoid duplicate worksheet titles .NET | C# iterate workbook worksheets Aspose.Cells | load and persist Excel workbook Aspose.Cells | error handling for missing Excel file C#

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetRenamerApp
{
    // The WorksheetRenamer class loads a workbook, iterates each worksheet, builds a name using the sheet's TabId (e.g., Sheet_3), ensures the name is unique within the workbook, assigns it to the worksheet, and saves the modified file to the specified output path.
    public class WorksheetRenamer
    {
        /// <param name="inputFilePath">Full path to the source Excel file.</param>
        /// <param name="outputFilePath">Full path where the modified Excel file will be saved.</param>
        public void RenameWorksheetsByTabId(string inputFilePath, string outputFilePath)
        {
            try
            {
                // Verify that the input file exists.
                if (!File.Exists(inputFilePath))
                    throw new FileNotFoundException($"Input file not found: {inputFilePath}");

                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputFilePath);

                // Iterate through all worksheets in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Retrieve the TabId of the current worksheet.
                    int tabId = sheet.TabId;

                    // Build a new name using the TabId.
                    string newName = $"Sheet_{tabId}";

                    // Ensure the new name is unique within the workbook.
                    int suffix = 1;
                    string uniqueName = newName;
                    while (workbook.Worksheets[uniqueName] != null)
                    {
                        uniqueName = $"{newName}_{suffix}";
                        suffix++;
                    }

                    // Rename the worksheet.
                    sheet.Name = uniqueName;
                }

                // Save the modified workbook to the specified output path.
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide input and output file paths.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            WorksheetRenamer renamer = new WorksheetRenamer();
            renamer.RenameWorksheetsByTabId(inputPath, outputPath);
        }
    }
}
