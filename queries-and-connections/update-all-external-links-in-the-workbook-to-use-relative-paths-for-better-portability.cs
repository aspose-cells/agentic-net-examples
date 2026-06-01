using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UpdateExternalLinksToRelativePaths
    {
        /// <summary>
        /// Loads a workbook, converts all absolute external link paths to relative paths,
        /// and saves the workbook.
        /// </summary>
        /// <param name="inputFile">Full path of the source workbook.</param>
        /// <param name="outputFile">Full path where the updated workbook will be saved.</param>
        public static void Run(string inputFile, string outputFile)
        {
            try
            {
                // Verify that the source workbook exists to avoid FileNotFoundException.
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Source file not found: {inputFile}");
                    return;
                }

                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputFile);

                // Base directory for relative path calculation – the folder containing the workbook.
                string baseDirectory = Path.GetDirectoryName(inputFile) ?? string.Empty;

                // Get the collection of external links.
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                // Iterate through each external link and convert its DataSource to a relative path.
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];
                    string currentPath = link.DataSource;

                    // Only process if the path is absolute (e.g., starts with a drive letter or UNC).
                    if (!string.IsNullOrEmpty(currentPath) && Path.IsPathRooted(currentPath))
                    {
                        // Compute the relative path from the workbook's folder to the external file.
                        string relativePath = Path.GetRelativePath(baseDirectory, currentPath);

                        // Update the external link to use the relative path.
                        link.DataSource = relativePath;
                    }
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the desired location.
                workbook.Save(outputFile);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            string sourcePath = @"C:\Data\MainWorkbook.xlsx";
            string destinationPath = @"C:\Data\MainWorkbook_Relative.xlsx";

            Run(sourcePath, destinationPath);

            Console.WriteLine("External links have been updated to relative paths and saved to:");
            Console.WriteLine(destinationPath);
        }
    }
}