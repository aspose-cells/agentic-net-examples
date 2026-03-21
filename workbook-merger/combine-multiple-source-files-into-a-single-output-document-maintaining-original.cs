using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CombineMultipleWorkbooksDemo
    {
        public static void Run()
        {
            // Paths of source workbooks to be combined
            string[] sourceFiles = new string[]
            {
                "Source1.xlsx",
                "Source2.xlsx",
                "Source3.xlsx"
            };

            // Ensure source files exist (for demo purposes we create simple workbooks if missing)
            for (int i = 0; i < sourceFiles.Length; i++)
            {
                if (!File.Exists(sourceFiles[i]))
                {
                    // Create a simple workbook with identifiable content
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets[0].Cells["A1"].PutValue($"Data from {Path.GetFileName(sourceFiles[i])}");
                    tempWb.Save(sourceFiles[i], SaveFormat.Xlsx);
                }
            }

            // Destination workbook – start with an empty workbook
            Workbook destWorkbook = new Workbook();

            // Combine each source workbook into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook
                Workbook srcWorkbook = new Workbook(filePath);

                // Merge the source workbook into the destination workbook
                destWorkbook.Combine(srcWorkbook);
            }

            // Save the combined workbook
            string outputPath = "CombinedOutput.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Combined workbook saved to: {outputPath}");
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