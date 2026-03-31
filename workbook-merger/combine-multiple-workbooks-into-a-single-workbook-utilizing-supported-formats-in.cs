using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public static class WorkbookCombiner
    {
        public static void CombineWorkbooks(string[] sourceFiles, string outputFile)
        {
            // Ensure source files exist; create empty workbooks if they don't.
            foreach (string filePath in sourceFiles)
            {
                if (!File.Exists(filePath))
                {
                    var tempWb = new Workbook();
                    tempWb.Save(filePath, SaveFormat.Xlsx);
                }
            }

            // Create an empty destination workbook.
            var destination = new Workbook(FileFormatType.Xlsx);

            // Merge each source workbook into the destination.
            foreach (string filePath in sourceFiles)
            {
                var source = new Workbook(filePath);
                destination.Combine(source);
            }

            // Save the combined workbook.
            destination.Save(outputFile, SaveFormat.Xlsx);
        }

        public static void Run()
        {
            string[] filesToMerge = new[]
            {
                "Report_Q1.xlsx",
                "Report_Q2.xlsx",
                "Report_Q3.xlsx",
                "Report_Q4.xlsx"
            };

            string combinedFile = "AnnualReport.xlsx";

            CombineWorkbooks(filesToMerge, combinedFile);

            Console.WriteLine($"Workbooks combined successfully into '{combinedFile}'.");
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