using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CombineWorksheetsDemo
    {
        public static void Run()
        {
            // Destination workbook – will contain all combined worksheets
            Workbook destWorkbook = new Workbook();

            // List of source workbook file paths to be merged
            string[] sourceFiles = new string[]
            {
                "Source1.xlsx",
                "Source2.xlsx",
                "Source3.xlsx"
            };

            // Iterate through each source file, load it, and combine into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook srcWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook
                destWorkbook.Combine(srcWorkbook);
            }

            // Save the combined workbook to disk
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"All worksheets combined successfully. Output saved to: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CombineWorksheetsDemo.Run();
        }
    }
}