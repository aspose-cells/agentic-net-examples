using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    public class WorkbookMerger
    {
        public static void MergeWorkbooks(string[] sourceFiles, string outputFile)
        {
            if (sourceFiles == null || sourceFiles.Length == 0)
                throw new ArgumentException("No source files provided.");

            using (Workbook destination = new Workbook())
            {
                foreach (string filePath in sourceFiles)
                {
                    if (!File.Exists(filePath))
                        throw new FileNotFoundException($"Source file not found: {filePath}");

                    using (Workbook source = new Workbook(filePath))
                    {
                        destination.Combine(source);
                    }
                }

                destination.Save(outputFile, SaveFormat.Xlsx);
            }
        }

        public static void Run()
        {
            string[] filesToMerge = new string[]
            {
                "Report_January.xlsx",
                "Report_February.xlsx",
                "Report_March.xlsx"
            };

            string mergedFile = "CombinedReport_Q1.xlsx";

            try
            {
                MergeWorkbooks(filesToMerge, mergedFile);
                Console.WriteLine($"Workbooks merged successfully. Output saved to: {mergedFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merging: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            WorkbookMerger.Run();
        }
    }
}