using System;
using Aspose.Cells;

namespace WorkbookConsolidationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example source files to be merged
            string[] sourceFiles = new string[]
            {
                "Report_January.xlsx",
                "Report_February.xlsx",
                "Report_March.xlsx"
            };

            // Destination file path
            string outputFile = "ConsolidatedReport.xlsx";

            // Perform consolidation
            ConsolidateWorkbooks(sourceFiles, outputFile);

            Console.WriteLine($"Workbooks have been consolidated into '{outputFile}'.");
        }

        /// <summary>
        /// Merges multiple Excel workbooks into a single workbook.
        /// Preserves worksheets, styles, data, and formulas.
        /// </summary>
        /// <param name="sourceFiles">Array of file paths to be merged.</param>
        /// <param name="outputFile">Path for the resulting merged workbook.</param>
        static void ConsolidateWorkbooks(string[] sourceFiles, string outputFile)
        {
            // Create an empty destination workbook (default format is Xlsx)
            Workbook destinationWorkbook = new Workbook();

            // Iterate through each source workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook
                // This merges worksheets, styles, data, and formulas.
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the consolidated workbook to the specified output path
            destinationWorkbook.Save(outputFile, SaveFormat.Xlsx);
        }
    }
}