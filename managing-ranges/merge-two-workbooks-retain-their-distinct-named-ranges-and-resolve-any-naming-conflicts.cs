using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MergeWorkbooksWithNamedRanges
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string targetPath = "TargetWorkbook.xlsx";
            const string sourcePath = "SourceWorkbook.xlsx";
            const string outputPath = "MergedWorkbook.xlsx";

            // Verify that the input files exist to avoid FileNotFoundException
            if (!File.Exists(targetPath))
                throw new FileNotFoundException($"Target workbook not found: {targetPath}");
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

            // Load the target workbook (will receive merged content)
            Workbook targetWorkbook = new Workbook(targetPath);

            // Load the source workbook (its sheets and named ranges will be merged)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Combine the source workbook into the target workbook (worksheets, charts, tables, etc.)
            targetWorkbook.Combine(sourceWorkbook);

            // Remove duplicate named ranges, keeping the first occurrence
            targetWorkbook.Worksheets.Names.RemoveDuplicateNames();

            // Optional: sort the remaining named ranges for better organization
            targetWorkbook.Worksheets.SortNames();

            // Save the merged workbook
            targetWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Merged workbook saved to: {outputPath}");
        }
    }
}