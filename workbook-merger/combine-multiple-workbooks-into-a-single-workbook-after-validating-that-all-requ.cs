using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CombineMultipleWorkbooksDemo
    {
        public static void Run()
        {
            // Paths of the workbooks that need to be merged.
            string[] sourceFiles = new string[] { "File1.xlsx", "File2.xlsx", "File3.xlsx" };

            // -----------------------------------------------------------------
            // Prerequisite validation: ensure every source file exists on disk.
            // -----------------------------------------------------------------
            foreach (string file in sourceFiles)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"Error: Source file not found -> {file}");
                    return; // Abort if any file is missing.
                }
            }

            try
            {
                // -------------------------------------------------------------
                // Load the first workbook – it will become the destination workbook.
                // -------------------------------------------------------------
                Workbook combinedWorkbook = new Workbook(sourceFiles[0]);

                // -------------------------------------------------------------
                // Iterate over the remaining workbooks and combine them one by one.
                // -------------------------------------------------------------
                for (int i = 1; i < sourceFiles.Length; i++)
                {
                    Workbook wbToCombine = new Workbook(sourceFiles[i]);
                    combinedWorkbook.Combine(wbToCombine);
                }

                // -------------------------------------------------------------
                // Save the resulting workbook to a new file.
                // -------------------------------------------------------------
                string outputPath = "CombinedWorkbook.xlsx";
                combinedWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"All workbooks merged successfully. Output saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during merging: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CombineMultipleWorkbooksDemo.Run();
        }
    }
}