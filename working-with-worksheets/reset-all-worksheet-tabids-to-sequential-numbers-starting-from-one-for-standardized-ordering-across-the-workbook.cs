// Title: Renumber worksheet TabId values sequentially in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set worksheet TabId sequentially starting at 1 | reset Excel sheet tab identifiers programmatically with Aspose.Cells | assign incremental TabId to each worksheet in a workbook using Aspose.Cells | batch update worksheet TabId property in C# Aspose.Cells example | change Excel worksheet tab order by modifying TabId with Aspose.Cells
// Tags: Aspose.Cells sequential TabId assignment | C# worksheet TabId reset | Excel workbook tab ordering Aspose | programmatic sheet TabId update .NET | modify worksheet TabId property

using System;
using System.IO;
using Aspose.Cells;

namespace ResetWorksheetTabIds
{
    // This C# example loads an Excel file with Aspose.Cells, iterates through all worksheets, assigns each sheet a TabId starting at 1 and incrementing sequentially, logs the changes, and saves the workbook to a new file (or overwrites the original).
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine input and output file paths
                string inputPath;
                string outputPath;

                if (args.Length >= 1 && File.Exists(args[0]))
                {
                    inputPath = args[0];
                }
                else
                {
                    Console.Write("Enter the full path to the Excel file to process: ");
                    inputPath = Console.ReadLine()?.Trim('"');

                    if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath))
                    {
                        Console.WriteLine("The specified file does not exist. Exiting.");
                        return;
                    }
                }

                // By default, overwrite the original file. User can specify a second argument for a different output.
                if (args.Length >= 2)
                {
                    outputPath = args[1];
                }
                else
                {
                    // Create a copy with suffix "_TabIdReset" to avoid accidental data loss
                    string directory = Path.GetDirectoryName(inputPath);
                    string filenameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                    string extension = Path.GetExtension(inputPath);
                    outputPath = Path.Combine(directory, $"{filenameWithoutExt}_TabIdReset{extension}");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Reset TabId for each worksheet sequentially starting from 1
                int nextTabId = 1;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.TabId = nextTabId;
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" assigned TabId = {nextTabId}");
                    nextTabId++;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with updated TabIds to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
