using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMergeApp
{
    public class WorkbookMergeService
    {
        /// <summary>
        /// Merges two Excel workbooks and returns the result as a MemoryStream.
        /// </summary>
        /// <param name="firstFilePath">Path to the destination workbook.</param>
        /// <param name="secondFilePath">Path to the source workbook.</param>
        /// <returns>MemoryStream containing the merged workbook (XLSX).</returns>
        public MemoryStream MergeWorkbooksToStream(string firstFilePath, string secondFilePath)
        {
            // Verify that both files exist before attempting to load them
            if (!File.Exists(firstFilePath))
                throw new FileNotFoundException($"Destination workbook not found: {firstFilePath}");
            if (!File.Exists(secondFilePath))
                throw new FileNotFoundException($"Source workbook not found: {secondFilePath}");

            try
            {
                // Load the destination and source workbooks
                Workbook destWorkbook = new Workbook(firstFilePath);
                Workbook sourceWorkbook = new Workbook(secondFilePath);

                // Merge source workbook into destination workbook
                destWorkbook.Combine(sourceWorkbook);

                // Save merged workbook into a memory stream
                MemoryStream mergedStream = new MemoryStream();
                destWorkbook.Save(mergedStream, SaveFormat.Xlsx);
                mergedStream.Position = 0; // Reset position for reading

                return mergedStream;
            }
            catch (Exception ex)
            {
                // Wrap any exception for clearer caller handling
                throw new InvalidOperationException("Failed to merge workbooks.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example file paths – replace with actual paths as needed
            string destPath = "dest.xlsx";
            string sourcePath = "source.xlsx";

            try
            {
                var service = new WorkbookMergeService();
                using (MemoryStream mergedStream = service.MergeWorkbooksToStream(destPath, sourcePath))
                {
                    // Save the merged workbook to a physical file for verification
                    string outputPath = "merged.xlsx";
                    using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        mergedStream.CopyTo(file);
                    }
                    Console.WriteLine($"Merged workbook saved to {outputPath}");
                }
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}