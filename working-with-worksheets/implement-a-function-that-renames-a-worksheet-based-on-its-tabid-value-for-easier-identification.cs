using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetRenameByTabId
    {
        /// <summary>
        /// Loads a workbook, renames each worksheet to include its TabId, and saves the result.
        /// </summary>
        /// <param name="inputPath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the modified workbook will be saved.</param>
        public static void RenameWorksheets(string inputPath, string outputPath)
        {
            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the collection
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Build a name that contains the TabId value
                    string proposedName = $"Sheet_{ws.TabId}";

                    // Ensure the name complies with Excel naming rules
                    string safeName = CellsHelper.CreateSafeSheetName(proposedName);

                    // Assign the safe name to the worksheet
                    ws.Name = safeName;
                }

                // Save the modified workbook to the specified output path
                workbook.Save(outputPath);
                Console.WriteLine($"Worksheets renamed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Example usage
        public static void Run()
        {
            string sourceFile = "input.xlsx";
            string resultFile = "output_renamed.xlsx";

            RenameWorksheets(sourceFile, resultFile);
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length >= 2)
                {
                    // Use command‑line arguments if provided
                    RenameWorksheets(args[0], args[1]);
                }
                else
                {
                    // Fallback to default example
                    Run();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}