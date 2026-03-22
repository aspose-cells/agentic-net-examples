using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    /// <summary>
    /// Demonstrates how to merge multiple Excel workbooks into a single workbook
    /// while preserving all worksheets, formulas, and formatting.
    /// </summary>
    public static class WorkbookCombiner
    {
        /// <summary>
        /// Combines the workbooks specified in <paramref name="inputFiles"/> into one workbook
        /// and saves the result to <paramref name="outputFile"/>.
        /// </summary>
        /// <param name="inputFiles">Array of full paths to the source workbooks.</param>
        /// <param name="outputFile">Full path where the combined workbook will be saved.</param>
        public static void CombineWorkbooks(string[] inputFiles, string outputFile)
        {
            if (inputFiles == null || inputFiles.Length == 0)
                throw new ArgumentException("At least one input file must be provided.", nameof(inputFiles));

            // Load the first workbook as the destination workbook.
            Workbook destWorkbook = new Workbook(inputFiles[0]);

            // Combine the remaining workbooks one by one.
            for (int i = 1; i < inputFiles.Length; i++)
            {
                Workbook sourceWorkbook = new Workbook(inputFiles[i]);
                destWorkbook.Combine(sourceWorkbook);
            }

            // Save the combined workbook to the specified output path.
            destWorkbook.Save(outputFile, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Run()
        {
            string[] filesToMerge = new string[]
            {
                "File1.xlsx",
                "File2.xlsx",
                "File3.xlsx"
            };

            string outputPath = "CombinedWorkbook.xlsx";

            CombineWorkbooks(filesToMerge, outputPath);

            Console.WriteLine($"Workbooks merged successfully. Output saved to: {outputPath}");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            WorkbookCombiner.Run();
        }
    }
}