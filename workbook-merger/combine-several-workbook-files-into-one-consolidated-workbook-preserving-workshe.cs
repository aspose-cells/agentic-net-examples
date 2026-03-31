using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConsolidateWorkbooks
    {
        /// <summary>
        /// Combines multiple Excel files into a single workbook.
        /// </summary>
        /// <param name="inputFiles">Array of file paths to be merged.</param>
        /// <param name="outputFile">Path of the consolidated workbook.</param>
        public static void Run(string[] inputFiles, string outputFile)
        {
            // Create the destination workbook (empty workbook)
            Workbook destinationWorkbook = new Workbook();

            // Iterate over each source file, load it, and combine into the destination
            foreach (string filePath in inputFiles)
            {
                string fullPath = Path.GetFullPath(filePath);

                if (!File.Exists(fullPath))
                {
                    // If the source file does not exist, create a placeholder workbook for demo purposes
                    Workbook placeholder = new Workbook();
                    placeholder.Worksheets[0].Name = Path.GetFileNameWithoutExtension(fullPath);
                    placeholder.Save(fullPath, SaveFormat.Xlsx);
                }

                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(fullPath);

                // Merge the source workbook into the destination workbook
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the consolidated workbook to the specified output path
            destinationWorkbook.Save(outputFile, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Main()
        {
            // Define the files to be merged
            string[] filesToMerge = new string[]
            {
                "Report_January.xlsx",
                "Report_February.xlsx",
                "Report_March.xlsx"
            };

            // Define the output file path
            string consolidatedFile = "Consolidated_Report_Q1.xlsx";

            // Perform the consolidation
            Run(filesToMerge, consolidatedFile);

            Console.WriteLine($"Workbooks merged successfully into '{consolidatedFile}'.");
        }
    }
}