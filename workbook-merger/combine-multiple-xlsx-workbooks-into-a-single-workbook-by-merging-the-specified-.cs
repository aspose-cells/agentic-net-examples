using System;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    public class WorkbookMerger
    {
        /// <summary>
        /// Merges multiple XLSX workbooks into a single workbook.
        /// </summary>
        /// <param name="sourceFiles">Paths of the source workbooks to merge.</param>
        /// <param name="outputFile">Path of the combined workbook to save.</param>
        public static void MergeWorkbooks(string[] sourceFiles, string outputFile)
        {
            // Create an empty destination workbook (default format is Xlsx)
            Workbook destWorkbook = new Workbook();

            // Iterate through each source file, load it, and combine into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook
                destWorkbook.Combine(sourceWorkbook);
            }

            // Save the combined workbook to the specified output path
            destWorkbook.Save(outputFile, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Run()
        {
            // Define source workbook file paths
            string[] filesToMerge = new string[]
            {
                "File1.xlsx",
                "File2.xlsx",
                "File3.xlsx"
            };

            // Define the output file path
            string combinedFile = "CombinedWorkbook.xlsx";

            // Perform the merge operation
            MergeWorkbooks(filesToMerge, combinedFile);

            Console.WriteLine($"Workbooks merged successfully into '{combinedFile}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookMerger.Run();
        }
    }
}