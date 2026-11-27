// Title: Validate and correct worksheet TabId values (1‑65535) using Aspose.Cells for .NET
// Common Searches: C# program to verify Excel worksheet TabId is between 1 and 65535 with Aspose.Cells | How to clamp out‑of‑range TabId values in a workbook using Aspose.Cells .NET | Command‑line tool for fixing invalid worksheet TabId numbers in Excel files via Aspose.Cells | Save Excel workbook only when worksheet TabId adjustments are made in Aspose.Cells | Ensuring legacy Excel compatibility by adjusting worksheet TabId in C#
// Tags: Aspose.Cells worksheet TabId validation | adjust Excel sheet TabId .NET | clamp TabId range Aspose.Cells | legacy Excel compatibility TabId correction | command‑line workbook TabId fixer

using System;
using System.IO;
using Aspose.Cells;

namespace ValidateTabId
{
    // A C# console application that loads an Excel workbook with Aspose.Cells, iterates through each worksheet, checks whether the TabId is outside the 1‑65535 range required for older Excel versions, clamps out‑of‑range values to the nearest valid limit, logs any adjustments, and saves the file only if changes were made or an explicit output path is provided, with full error handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args == null || args.Length == 0)
                {
                    Console.WriteLine("Error: No input file specified.");
                    PrintUsage();
                    return;
                }

                string inputPath = args[0];
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }

                string outputPath;
                if (args.Length >= 2 && !string.IsNullOrWhiteSpace(args[1]))
                {
                    outputPath = args[1];
                }
                else
                {
                    // Overwrite the original file after validation
                    outputPath = inputPath;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                bool anyChanges = false;

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int originalTabId = sheet.TabId;
                    if (originalTabId < 1 || originalTabId > 65535)
                    {
                        // Clamp the value to the valid range
                        int correctedTabId = originalTabId < 1 ? 1 : 65535;
                        sheet.TabId = correctedTabId;
                        anyChanges = true;
                        Console.WriteLine($"Worksheet \"{sheet.Name}\": TabId {originalTabId} is out of range. Adjusted to {correctedTabId}.");
                    }
                    else
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\": TabId {originalTabId} is within the valid range.");
                    }
                }

                // Save the workbook if any changes were made or if the user explicitly requested an output file
                if (anyChanges || !string.Equals(inputPath, outputPath, StringComparison.OrdinalIgnoreCase))
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to \"{outputPath}\".");
                }
                else
                {
                    Console.WriteLine("No changes required. Workbook not saved.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Prints command‑line usage information.
        /// </summary>
        private static void PrintUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  ValidateTabId.exe <inputWorkbookPath> [outputWorkbookPath]");
            Console.WriteLine();
            Console.WriteLine("Parameters:");
            Console.WriteLine("  <inputWorkbookPath>   Path to the Excel workbook to validate.");
            Console.WriteLine("  [outputWorkbookPath]  Optional path to save the validated workbook.");
            Console.WriteLine("                         If omitted, the input file will be overwritten.");
        }
    }
}
