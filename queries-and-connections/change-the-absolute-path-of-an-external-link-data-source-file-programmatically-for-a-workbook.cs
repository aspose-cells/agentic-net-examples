// Title: Programmatically Update Absolute Paths of External Links in an Aspose.Cells Workbook (C#)
// Description: Loads a workbook, iterates through its Worksheets.ExternalLinks collection, replaces a specified folder segment in each link's OriginalDataSource, writes the new path back, and saves the file. Includes checks for missing input files and automatic creation of the output directory.
// Keywords: Aspose.Cells external links | C# update link path | change OriginalDataSource | modify absolute path Excel workbook | .NET external link path replacement | batch fix broken links Aspose
// Common Searches: how to change external link path Aspose.Cells C# | update OriginalDataSource for all links in workbook | replace folder segment in Excel external links programmatically | Aspose.Cells move source files and fix links | C# code to edit external link paths in Excel
// Developer Intent: Replace the stored absolute file paths of every external link in a workbook with a new folder location using Aspose.Cells for .NET.
// Use Cases: Migrate workbooks after relocating source data to a different directory. | Deploy Excel files to a new server environment where linked files reside in another folder. | Run a batch job that repairs broken external references after a folder restructuring.
// AI Prompts: Write C# code with Aspose.Cells that substitutes an old folder path with a new one in the OriginalDataSource of all external links and saves the workbook. | Explain how to confirm that external link paths were updated correctly after modifying OriginalDataSource. | Create a reusable method that takes oldFolder and newFolder strings and updates every external link path in a given Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, iterates through its Worksheets.ExternalLinks collection, replaces a specified folder segment in each link's OriginalDataSource, writes the new path back, and saves the file. Includes checks for missing input files and automatic creation of the output directory.
    public class ChangeExternalLinkPath
    {
        public static void Run()
        {
            try
            {
                // Load the workbook that contains external links
                string inputPath = @"C:\Input\WorkbookWithLinks.xlsx";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                Workbook workbook = new Workbook(inputPath);

                // Define the part of the path to replace and the new replacement
                string oldPathPart = @"C:\OldFolder\";
                string newPathPart = @"D:\NewFolder\";

                // Iterate through all external links and modify their stored data source paths
                for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
                {
                    // Get the original stored data source (absolute path)
                    string original = workbook.Worksheets.ExternalLinks[i].OriginalDataSource;

                    // Replace the old part with the new part
                    string updated = original.Replace(oldPathPart, newPathPart);

                    // Assign the modified path back to the external link
                    workbook.Worksheets.ExternalLinks[i].OriginalDataSource = updated;
                }

                // Save the workbook with the updated external link paths
                string outputPath = @"C:\Output\WorkbookWithLinks_Updated.xlsx";

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ChangeExternalLinkPath.Run();
        }
    }
}
