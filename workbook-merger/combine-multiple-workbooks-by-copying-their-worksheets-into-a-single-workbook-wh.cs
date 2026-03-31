using System;
using Aspose.Cells;

namespace WorkbookCombinerDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Paths of workbooks to be combined
            string[] sourceFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Destination file path
            string outputFile = "CombinedWorkbook.xlsx";

            // Combine the workbooks
            CombineWorkbooks(sourceFiles, outputFile);

            Console.WriteLine($"Workbooks combined successfully. Output saved to '{outputFile}'.");
        }

        /// <summary>
        /// Combines multiple Excel workbooks into a single workbook.
        /// All worksheets, cell data and formatting are preserved.
        /// </summary>
        /// <param name="sourceFiles">Array of file paths to source workbooks.</param>
        /// <param name="outputFile">Path for the combined workbook.</param>
        public static void CombineWorkbooks(string[] sourceFiles, string outputFile)
        {
            // Create an empty destination workbook.
            // This uses the provided constructor rule.
            Workbook destinationWorkbook = new Workbook();

            // Optional: remove the default empty worksheet if you don't want it in the final file.
            // It will be removed after the first combine if it remains empty.
            if (destinationWorkbook.Worksheets.Count == 1 &&
                destinationWorkbook.Worksheets[0].Cells.MaxDataRow == -1 &&
                destinationWorkbook.Worksheets[0].Cells.MaxDataColumn == -1)
            {
                destinationWorkbook.Worksheets.RemoveAt(0);
            }

            // Iterate through each source workbook file.
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file.
                // This uses the provided constructor rule that accepts a file path.
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook.
                // The Combine method merges all worksheets, preserving data and formatting.
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the combined workbook to the specified output file.
            // This uses the provided Save method with format specification.
            destinationWorkbook.Save(outputFile, SaveFormat.Xlsx);
        }
    }
}