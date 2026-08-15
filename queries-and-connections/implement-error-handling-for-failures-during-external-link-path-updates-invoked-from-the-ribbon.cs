// Title: C# – Robust error handling for updating external link paths via Aspose.Cells ribbon command
// Description: Loads an Excel workbook, iterates through its ExternalLinkCollection, replaces the folder segment of each link's DataSource, assigns the new path, and saves the file. Each step (load, per‑link update, save) is wrapped in try‑catch blocks that log errors and allow processing to continue when invoked from a ribbon UI.
// Keywords: Aspose.Cells external link update | C# error handling Excel links | replace DataSource path Aspose | ribbon command workbook save | exception handling external links
// Common Searches: how to handle exceptions when updating external links in Aspose.Cells | C# replace folder in external link DataSource with try‑catch | save workbook after modifying external links Aspose.Cells | ribbon button update external link paths error handling
// Developer Intent: Add layered try‑catch logic to safely modify external link paths and persist the workbook without crashing the ribbon UI.
// Use Cases: Update all external link paths in a workbook while skipping links that cause errors. | Log detailed messages for each failed link update and still save the modified file. | Provide a resilient ribbon‑triggered routine that handles load, per‑link, and save failures separately.
// AI Prompts: Write C# code that loops through workbook.Worksheets.ExternalLinks, replaces a base folder in each DataSource, and logs index‑specific exceptions. | Show a pattern for loading an Excel file, updating external link paths, and saving it with distinct try‑catch blocks suitable for a ribbon command. | Suggest ways to display per‑link update errors in a UI dialog or write them to a log file when the path replacement fails.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkUpdate
{
    // Loads an Excel workbook, iterates through its ExternalLinkCollection, replaces the folder segment of each link's DataSource, assigns the new path, and saves the file. Each step (load, per‑link update, save) is wrapped in try‑catch blocks that log errors and allow processing to continue when invoked from a ribbon UI.
    public class UpdateExternalLinkPaths
    {
        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through all external links and update their paths
            for (int i = 0; i < externalLinks.Count; i++)
            {
                try
                {
                    // Example: replace an old base folder with a new one
                    string oldPath = externalLinks[i].DataSource;
                    string newPath = oldPath.Replace(@"C:\OldFolder\", @"D:\NewFolder\");

                    // Assign the new path
                    externalLinks[i].DataSource = newPath;
                    externalLinks[i].OriginalDataSource = newPath;

                    Console.WriteLine($"Link {i} updated successfully: {newPath}");
                }
                catch (Exception ex)
                {
                    // Log the error but continue processing remaining links
                    Console.WriteLine($"Error updating link at index {i}: {ex.Message}");
                }
            }

            try
            {
                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                UpdateExternalLinkPaths.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
